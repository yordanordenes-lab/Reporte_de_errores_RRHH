# Auditoría RRHH - Asistencia vs Liquidación

Script en Python experimental para automatizar la revisión cruzada entre registros de asistencia y liquidaciones de sueldo en PDF (formato chileno)

⚠️ **Proyecto educativo con datos 100% ficticios.**  
- Uso educativo y experimental.
- Los datos de prueba incluidos son ficticios.
- No procesar datos reales sin salvaguardas legales (Ley 19.628, GDPR).

## Requisitos
- Python 3.8+
- Las dependencias del archivo `requirements.txt`

```bash
pip install -r requirements.txt
```

## Uso básico

```bash
# Modo automático (detecta PDFs por nombre en la carpeta)
python auditoria_mensual.py --carpeta ./datos/

# Modo manual (rutas específicas)
python auditoria_mensual.py \
  --asistencia ./datos/asistencia.pdf \
  --liquidacion ./datos/liquidacion.pdf \
  --salida reporte.xlsx
```

## Salida

Genera un archivo Excel con tres hojas:

- Hallazgos: discrepancias detectadas
- Duplicados: RUTs repetidos
- Resumen: métricas generales
