# Operación

## Variables de entorno (PowerShell, antes de arrancar Spring)

```powershell
$env:AZURE_TENANT_ID = "<tenant-guid>"
$env:AZURE_CLIENT_ID = "<client-id>"
$env:AZURE_CLIENT_SECRET = "<client-secret>"
```

## Terminal 1 — Spring Boot (VS Code o PowerShell)

```powershell
cd dynamics-integration
.\run.ps1
```

Verificación:

```powershell
Invoke-RestMethod http://localhost:8080/actuator/health
Invoke-RestMethod http://localhost:8080/api/token/refresh -Method Post | ConvertTo-Json
Invoke-RestMethod http://localhost:8080/api/health/dynamics | ConvertTo-Json
```

Respuesta esperada de health: `"success": true`.

## Terminal 2 — ngrok

```powershell
ngrok http 8080
```

Confirme en consola:

```text
https://<dominio>.ngrok-free.dev -> http://localhost:8080
```

Prueba en navegador:

```text
https://<dominio>.ngrok-free.dev/api/health/dynamics
```

## Excel Online

### Celda de URL

| Celda | Valor |
|-------|--------|
| **Resultado!B1** (o A2) | `https://<dominio>.ngrok-free.dev` |

Sin barra final. Guarde el libro.

### Scripts (Automatizar)

Instale desde `office-scripts/dist/` o genere antes:

```powershell
cd office-scripts
.\build-office-scripts.ps1
```

| Script en Excel | Archivo |
|-----------------|---------|
| Probar Conexion | `dist/ProbarConexion.osts.ts` |
| Crear Pedido | `dist/CrearPedido.osts.ts` |
| Crear Lineas | `dist/CrearLineas.osts.ts` |

### Orden de ejecución

1. **Probar Conexion** → Estado `CONEXION OK`
2. **Crear Pedido** → Estado `CABECERA OK` + OV en **Pedido Dynamics**
3. **Crear Lineas** → Estado `LINEAS OK`

## Prueba local sin Excel

```powershell
cd dynamics-integration
.\test-api.ps1
```
