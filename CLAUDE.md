# CLAUDE.md — RUBEN-COTON_EXTENSION-EMAILING

## Que es este proyecto

Extension (add-on) de Google Sheets para la cuenta manager@rubencoton.com.
Conecta con el ecosistema Google completo: Sheets, Drive, Gmail, Calendar, Contacts.

## Estructura

```
RUBEN-COTON_EXTENSION-EMAILING/
├── .clasp.json              # Config clasp → apunta a src-addon/
├── src-addon/               # Codigo Apps Script (se despliega con clasp push)
│   ├── appsscript.json      # Manifest: scopes, triggers, add-on config
│   ├── Main.js              # Entry points: onOpen, onHomepage, sidebar
│   ├── Ecosystem.js         # Conexion con Sheets, Drive, Gmail, Calendar, Contacts
│   ├── CardActions.js       # Acciones de los botones del Card UI
│   └── Sidebar.html         # Panel lateral HTML
├── src/                     # Codigo Node.js local (utilidades, sync offline)
│   ├── sheetsConnector.js   # Conector Google Sheets (3 modos auth)
│   ├── listSheets.js        # npm run list-sheets
│   └── syncOnce.js          # npm run sync
├── docs/
│   ├── CREDENCIALES.md      # Como configurar credenciales
│   └── SHEETS_MAP.md        # Mapa de hojas conectadas
├── .env.example             # Template variables de entorno (Node.js)
└── package.json
```

## Dos entornos

| Entorno | Carpeta | Lenguaje | Deploy |
|---------|---------|----------|--------|
| Apps Script (add-on) | `src-addon/` | GS/JS (V8) | `npx clasp push` |
| Node.js (local) | `src/` | Node.js 20+ | `npm run sync` |

## Workflow de desarrollo

```bash
# Editar codigo del add-on
# Archivos en src-addon/ → se editan localmente

# Subir a Apps Script
npx clasp push

# Bajar cambios si se edito desde el editor web
npx clasp pull

# Abrir editor web
npx clasp open

# Ver logs
npx clasp logs

# Probar Node.js local
npm run list-sheets
npm run sync
```

## Reglas de desarrollo

1. **Archivos add-on usan .js** (no .gs). Clasp los convierte automaticamente.
2. **No editar appsscript.json desde el editor web** — siempre desde local y clasp push.
3. **Scopes**: si agregas un servicio nuevo, anadir el scope en appsscript.json.
4. **No subir secrets**: .env, token.json, secrets/ estan en .gitignore.
5. **CRM_SHEETS y ECOSYSTEM_SHEETS** en Ecosystem.js son la fuente de verdad de IDs.
6. **Solo lectura** sobre hojas CRM externas. No escribir en ellas.
7. **clasp push antes de commit**: asegurar que Apps Script esta sincronizado.

## Credenciales

- Cuenta: manager@rubencoton.com
- ScriptId: `1P5NO7A5X0PHpPfnTEOkQkXBbFspRuV8Osi4dmxDRmFaEyi-Yj9AQUjpJ`
- OAuth client: el mismo que usa todo el ecosistema (codex-google-hub)
- Clasp login: `npx clasp login` si el token expira

## Publicar como add-on

Aun no publicado. Para testear:
1. Abrir editor: `npx clasp open`
2. Deploy > Test deployments > Install
3. Abrir cualquier hoja de calculo → menu Extensiones → RUBEN COTON Emailing

## Git

Si `git push` directo falla, usar:
```
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\elrub\Desktop\CARPETA CODEX\03_SCRIPTS_UTILIDAD\publicar_desde_local.ps1" -RepoPath "C:\Users\elrub\Desktop\CARPETA CODEX\01_PROYECTOS\RUBEN-COTON_EXTENSION-EMAILING" -Remote origin -Branch main
```
