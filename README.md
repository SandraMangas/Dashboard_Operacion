# Dashboard de Operacion

Dashboard operativo construido con Streamlit, Pandas y Plotly.

## Estructura

- `app.py`: punto de entrada de la aplicacion Streamlit.
- `src/data/`: lectura y validacion de datos de entrada.
- `src/domain/`: reglas de negocio.
- `src/metrics/`: calculo de KPIs y metricas.
- `src/visualizations/`: construccion de graficos Plotly.
- `src/ui/`: componentes y composicion de la interfaz.
- `data/`: archivos locales de entrada, excluidos del control de versiones.

## Ejecucion

1. Crear un entorno virtual.
2. Instalar las dependencias de `requirements.txt`.
3. Colocar el archivo Excel en `data/` cuando se defina su ubicacion y estructura.
4. Ejecutar `streamlit run app.py`.
