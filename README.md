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

💡 Nota sobre los archivos de prueba: Al clonar el repositorio, los archivos LIQUIDACIÓN DE SUELDO.pdf y REGISTRO DE ASISTENCIA.pdf ya vienen incluidos automáticamente en la misma carpeta que el script.

Para ejecutar la auditoría, usa uno de los siguientes métodos según cómo abras el proyecto:

Opción A: Desde Visual Studio Code (Recomendado)
Si abres la carpeta del proyecto directamente en VS Code (File > Open Folder...), tu terminal integrada ya se iniciará en el lugar correcto. Solo debes ejecutar:

```bash
python Reporte_de_errores_RRHH.py
```
Opción B: Desde una terminal externa (PowerShell / CMD / Terminal)
Si ejecutas los comandos desde una consola limpia por separado, la terminal iniciará "afuera" del proyecto. Debes usar el comando cd para entrar a la carpeta antes de despertar a Python:

```bash
# 1. Entra a la carpeta del proyecto clonado
cd Reporte_de_errores_RRHH

# 2. Ejecuta el script
python Reporte_de_errores_RRHH.py
```

## Salida

Genera un archivo Excel con tres hojas:

- Hallazgos: discrepancias detectadas
- Duplicados: RUTs repetidos
- Resumen: métricas generales
