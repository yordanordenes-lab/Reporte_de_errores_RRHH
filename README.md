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

Para ejecutar el script y comenzar la auditoría automatizada, simplemente abre la terminal en la raíz del proyecto y ejecuta el siguiente comando:

```bash
python Reporte_de_errores_RRHH.py
```

## Salida

Genera un archivo Excel con tres hojas:

- Hallazgos: discrepancias detectadas
- Duplicados: RUTs repetidos
- Resumen: métricas generales
