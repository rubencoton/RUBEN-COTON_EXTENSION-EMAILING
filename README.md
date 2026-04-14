# RUBEN-COTON_EXTENSION-EMAILING

Extension de emailing para RUBEN COTON.
Conecta con el ecosistema de Google Sheets de `manager@rubencoton.com`.

## Estado

- Fase: inicial
- Creado: 2026-04-14
- Cuenta Google: manager@rubencoton.com

## Estructura

```
RUBEN-COTON_EXTENSION-EMAILING/
├── src/
│   ├── sheetsConnector.js    # Conector Google Sheets (3 modos auth)
│   ├── listSheets.js         # Listar hojas conectadas
│   └── syncOnce.js           # Sync manual de emails
├── docs/
│   ├── CREDENCIALES.md       # Como configurar credenciales
│   └── SHEETS_MAP.md         # Mapa de todas las hojas conectadas
├── .env.example              # Template de variables de entorno
├── .gitignore
└── package.json
```

## Hojas conectadas

6 CRMs de Google Sheets con datos de contactos (solo lectura):

| Hoja | Tabs | Emails |
|------|------|--------|
| CRM: VENTA-BOOKING | 23 | EMAIL |
| CRM: OTROS | 12 | EMAIL |
| CRM: MUNDO DISCOGRAFICO | 4 | EMAIL |
| CRM: MARKETING Y PROMOCION | 9 | CORREO |
| CRM: FESTIVALES | 31 | EMAIL |
| CRM: BELLA BESTIA | 5 | EMAIL |

Ver detalle completo en [docs/SHEETS_MAP.md](docs/SHEETS_MAP.md).

## Arranque rapido

```bash
# 1. Instalar dependencias
npm install

# 2. Configurar credenciales
cp .env.example .env
# Editar .env con datos de manager@rubencoton.com
# Ver docs/CREDENCIALES.md para opciones

# 3. Verificar conexion
npm run list-sheets

# 4. Sync manual
npm run sync
```

## Conexion con otros proyectos

- **APP_ARTES-BUHO_EMAILING** — Plataforma de email marketing (Node.js/Express). Este proyecto extiende su ecosistema de hojas.
- **codex-google-hub** — Almacen central de credenciales OAuth.
- **drive-manager-rubencoton-com** — Credenciales OAuth reutilizables.
- **RUBEN-COTON_GMAIL** — Router Gmail con hoja CRM.
- **RUBEN-COTON_CALENDAR** — Calendar hub con hoja vinculada.
