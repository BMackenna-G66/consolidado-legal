# Solicitud a TI — Registro de aplicación en Azure AD

> Copiar/pegar como ticket o correo a TI. Trámite estimado: 10 minutos.

**Asunto:** Registro de app SPA en Azure AD para lectura de SharePoint (Consolidado Legal — Compliance)

Hola,

Necesito registrar una aplicación interna en Azure AD para que una herramienta de Compliance pueda leer, con la sesión del propio usuario, una carpeta de SharePoint a la que ese usuario ya tiene acceso. No requiere permisos de aplicación ni acceso a datos de otros usuarios.

**Qué hace la herramienta:** consolida los gastos legales por país leyendo los documentos que los encargados suben a la carpeta `Consolidado Cobros - Pagos [Compliance]`. Todo el procesamiento ocurre en el navegador del usuario; no hay backend ni almacenamiento externo.

**Lo que necesito que configuren:**

| Campo | Valor |
|---|---|
| Nombre | Consolidado Legal Compliance |
| Supported account types | Accounts in this organizational directory only (Single tenant) |
| Platform | **Single-page application (SPA)** |
| Redirect URI | `http://localhost:4173` (uso local actual) — y la URL definitiva cuando se publique |
| API permission | Microsoft Graph → **Delegated** → `Files.Read.All` |
| Client secret | **No se requiere** (SPA con PKCE) |

**Lo que necesito de vuelta:** el **Application (client) ID** del registro.

Notas:
- El permiso es *delegated*: la app actúa siempre como el usuario que inició sesión y solo ve lo que ese usuario ya puede ver. Si la política exige *admin consent* para `Files.Read.All`, favor otorgarlo.
- Si prefieren restringirlo más, `Files.Read.Selected` o `Sites.Selected` también sirven acotando a este sitio; avísenme y ajusto la configuración.

Gracias,
Benjamín Mackenna — Compliance
