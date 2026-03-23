# COA-Generator

Herramienta interna en Python/Tkinter para automatizar la generación de certificados de análisis (COA) a partir de packing lists PDF, plantillas `.docx` e ingreso de resultados microbiológicos.

## Persistencia y compatibilidad

El proyecto mantiene la compatibilidad con los archivos históricos existentes:

- `Historial_Microbiologia.xlsx`: historial microbiológico por formato.
- `Registro_COAs.xlsx`: registro anual de COAs generados.
- `session.json`: sesión de trabajo restaurable.
- `config.json`: configuración editable de rutas, formatos y mapeos.

La modularización actual separa esos accesos en `coa_storage.py` sin cambiar los nombres ni la estructura de los archivos, para reducir riesgo de pérdida del historial.

## Estructura actual

- `Generador_COA.py`: interfaz principal y lógica de generación.
- `coa_formats.py`: formatos y mapeos por defecto.
- `coa_storage.py`: carga/guardado de configuración, historial, registro y sesión.
