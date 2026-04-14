# Credenciales — RUBEN-COTON_EXTENSION-EMAILING

Cuenta principal: manager@rubencoton.com

## Como configurar

### Opcion A: OAuth (recomendada para hojas privadas)

Copiar desde el almacen central de credenciales:

```
Fuente: codex-google-hub/secrets/
```

1. `GOOGLE_CLIENT_ID` y `GOOGLE_CLIENT_SECRET` → desde `client_secret_from_clasprc_manager_rubencoton_com.json`
2. `GOOGLE_REFRESH_TOKEN` → desde `oauth_token_manager_rubencoton_com.json` (campo `refresh_token`)

Alternativa: copiar `.env` y `config/token.json` desde `drive-manager-rubencoton-com/`

### Opcion B: API Key (solo hojas publicas)

Usar `GOOGLE_SHEETS_API_KEY` si las hojas estan compartidas con "cualquiera con el enlace".

### Opcion C: Service Account

Pegar el JSON completo en `GOOGLE_SHEETS_CREDENTIALS` (una linea).

## Proyectos hermanos que usan la misma cuenta

| Proyecto | Servicios |
|----------|-----------|
| drive-manager-rubencoton-com | Drive, Gmail, Calendar |
| RUBEN-COTON_COPIA-CRM-ARTES-BUHO | Drive, Sheets |
| RUBEN-COTON_GMAIL | Gmail |
| RUBEN-COTON_CALENDAR | Calendar |
| APP_ARTES-BUHO_EMAILING | Sheets (solo lectura, 6 CRMs) |

## Seguridad

- NUNCA subir `.env`, `token.json` ni `secrets/` a GitHub
- Si se filtra un token, regenerar en Google Cloud Console
- El `.gitignore` ya excluye estos archivos
