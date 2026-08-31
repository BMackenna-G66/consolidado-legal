// Clasificación de partidas en Gastos / Honorarios / Juicios y otros.
// Prioridad: 1) override manual del mantenedor  2) clasificación IA (cacheada)
//            3) regla heurística  4) lo que sugirió el parser.
// La base histórica conserva el concepto ya reportado salvo override manual.

const KEY_MANUAL = 'consolidado-clasif-manual';
const KEY_IA = 'consolidado-clasif-ia';
const KEY_GEMINI = 'GEMINI_API_KEY';

export const CONCEPTOS = ['Gastos', 'Honorarios', 'Juicios y otros'];

const leer = k => { try { return JSON.parse(localStorage.getItem(k) || '{}'); } catch { return {}; } };
const guardar = (k, v) => localStorage.setItem(k, JSON.stringify(v));

export function normalizaCat(cat) {
  return String(cat || 'Sin categoría').trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

export function overrides() { return leer(KEY_MANUAL); }
export function setOverride(cat, concepto) {
  const m = leer(KEY_MANUAL);
  if (concepto) m[normalizaCat(cat)] = concepto; else delete m[normalizaCat(cat)];
  guardar(KEY_MANUAL, m);
}
export function cacheIA() { return leer(KEY_IA); }

function heuristica(cat) {
  const c = normalizaCat(cat);
  if (/honorario|asesoria legal|asesoria profesional|minuta|boleta|patrocinio|defensa|hora/.test(c)) return 'Honorarios';
  if (/demanda|querella|juicio|litigio|denuncia penal|proceso judicial|investigacion/.test(c)) return 'Juicios y otros';
  return 'Gastos';
}

// Devuelve { concepto, origen } para un registro ya parseado.
export function clasificar(registro) {
  const manual = leer(KEY_MANUAL)[normalizaCat(registro.categoria)];
  if (manual) return { concepto: manual, origen: 'Manual' };
  if (registro.fuente === 'base') return { concepto: registro.concepto, origen: 'Reportado' };
  // Parte prorrateada de una boleta de honorarios: es un pago real al estudio,
  // aunque el trabajo haya sido sobre juicios — no confundir con las cuantías
  // de "Juicios y otros".
  if (registro.prorrateo) return { concepto: registro.concepto || 'Honorarios', origen: 'Documento' };
  const ia = leer(KEY_IA)[normalizaCat(registro.categoria)];
  if (ia) return { concepto: ia, origen: 'IA' };
  const h = heuristica(registro.categoria);
  if (h !== 'Gastos') return { concepto: h, origen: 'Regla' };
  return { concepto: registro.concepto || 'Gastos', origen: 'Regla' };
}

// ---------- clasificación con IA (Gemini, mismo proveedor que SmartCheck) ----------

export function apiKeyGemini() { return sessionStorage.getItem(KEY_GEMINI) || ''; }
export function setApiKeyGemini(k) { sessionStorage.setItem(KEY_GEMINI, k.trim()); }

export async function clasificarConIA(categorias) {
  const key = apiKeyGemini();
  if (!key) throw new Error('Falta la API key de Gemini');
  const cache = leer(KEY_IA);
  const pendientes = [...new Set(categorias.map(normalizaCat))].filter(c => !cache[c]);
  if (!pendientes.length) return { nuevas: 0 };

  const prompt = `Eres analista de compliance de una fintech. Clasifica cada partida de gasto legal en exactamente una de estas clases:
- "Honorarios": pago por servicios profesionales de abogados o estudios jurídicos (asesorías, patrocinio, defensa, horas facturadas, boletas de honorarios, hitos de honorarios).
- "Gastos": gastos operativos o reembolsables (notaría, legalizaciones, movilidad, envíos, tasas, impuestos, gastos administrativos).
- "Juicios y otros": cuantías de demandas, querellas, litigios o investigaciones (montos en disputa, no pagos efectivos).

Responde SOLO un objeto JSON que mapee cada partida (texto exacto de entrada) a su clase.
Partidas: ${JSON.stringify(pendientes)}`;

  const r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(key)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: 'application/json', temperature: 0 },
    }),
  });
  if (!r.ok) throw new Error(`Gemini respondió ${r.status}`);
  const data = await r.json();
  const texto = data.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
  let mapa;
  try { mapa = JSON.parse(texto); } catch { throw new Error('La IA no devolvió JSON válido'); }
  let nuevas = 0;
  for (const [cat, cls] of Object.entries(mapa)) {
    if (CONCEPTOS.includes(cls)) { cache[normalizaCat(cat)] = cls; nuevas++; }
  }
  guardar(KEY_IA, cache);
  return { nuevas };
}
