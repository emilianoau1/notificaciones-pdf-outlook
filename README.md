# Notificaciones automáticas con PDF adjunto por registro

Script en Python que automatiza el envío de correos individuales desde Outlook,
buscando y adjuntando automáticamente el PDF correspondiente a cada cliente o unidad.

Incluye modo borrador para revisar antes de enviar, manejo de errores por registro
y generación automática de un reporte del proceso.

---

## Problema que resuelve

Cuando se necesita notificar a decenas o cientos de clientes con un documento PDF
diferente para cada uno, el proceso manual implica:

- Buscar el PDF de cada cliente en una carpeta
- Redactar y enviar cada correo por separado
- Llevar un registro de quién fue notificado y quién no

Este script hace todo eso automáticamente en minutos.

---

## Qué hace el script

1. Lee un Excel con los datos de cada cliente (ID, grupo, correo)
2. Por cada registro, busca en una carpeta el PDF que corresponde a ese ID
3. Crea el correo en Outlook con el PDF adjunto
4. Envía o guarda como borrador según la configuración
5. Si no encuentra el PDF, guarda el borrador para adjuntar manualmente
6. Al finalizar, genera un reporte Excel con el estado de cada envío

---

## Características destacadas

- Modo borrador (`SEND_DIRECTLY = False`) para revisar correos antes de enviar
- Búsqueda de PDF flexible: nombre exacto o sin distinción de mayúsculas
- Manejo de errores por registro: un fallo no detiene el proceso completo
- Reporte automático al final con estado de cada correo (enviado / borrador / error)
- Validación de columnas al inicio para detectar problemas antes de procesar

---

## Tecnologías

| Librería | Uso |
|---|---|
| `pandas` | Lectura del Excel y generación del reporte |
| `pathlib` | Búsqueda de archivos PDF en carpeta |
| `win32com` | Integración con Microsoft Outlook |
| `traceback` | Registro detallado de errores |

---

## Estructura del Excel fuente

| Columna | Descripción |
|---|---|
| `ID_UNIDAD` | Identificador único — debe coincidir con el nombre del PDF |
| `GRUPO_CLIENTE` | Nombre del grupo o empresa cliente |
| `EJECUTIVO` | Nombre del ejecutivo de cuenta |
| `CORREO` | Email del destinatario |

Los PDFs en la carpeta deben llamarse exactamente igual que el valor en `ID_UNIDAD`.
Ejemplo: si `ID_UNIDAD` es `ABC123`, el archivo debe ser `ABC123.pdf`.

---

## Cómo usarlo

### 1. Instalar dependencias

```bash
pip install pandas openpyxl pywin32
```

### 2. Configurar rutas

```python
EXCEL_PATH    = r'C:\ruta\a\CLIENTES.xlsx'
PDF_FOLDER    = r'C:\ruta\a\carpeta\PDFS'
CC_EMAILS     = "supervisor@empresa.com"
SEND_DIRECTLY = False  # Cambiar a True cuando estés listo para enviar
```

### 3. Ejecutar

```bash
python notificaciones_con_pdf.py
```

### Ejemplo de salida en consola

```
Leyendo Excel...
Se procesarán 45 registros...

[1/45] BORRADOR  — ABC123 → cliente1@empresa.com
[2/45] BORRADOR  — DEF456 → cliente2@empresa.com
[3/45] SIN PDF   — GHI789 (borrador guardado)
...

─────────────────────────────
Proceso finalizado.
  Total procesados : 45
  Enviados         : 0
  Borradores       : 44
  Errores          : 1
  Reporte guardado : C:\ruta\a\reporte_envios.xlsx
─────────────────────────────
```

---

## Casos de uso similares

- Envío de contratos, facturas o recibos individuales
- Notificaciones de cobranza con estado de cuenta adjunto
- Distribución de reportes personalizados por cliente
- Cualquier proceso donde cada destinatario recibe un PDF diferente

---

## Autor

Desarrollado como parte de una automatización real para gestión de contratos
en empresa del sector financiero-automotriz. Adaptable a cualquier industria.
