# Dashboard de Análisis de Rendimiento — Planta Carbonato de Litio

Dashboard interactivo en Streamlit que replica el análisis de la presentación v3.
Sube el Excel de planta y obtén toda la comparación entre periodos sobre meta y
bajo meta automáticamente.

## Qué incluye

- **7 pestañas de análisis**:
  1. Enfoque (control vs desempeño) + mapa global de correlaciones
  2. E1 — Control (flujos individuales, T AFT, tiempo de residencia real)
  3. E1 — Desempeño (KPI Li Bfil y Mg/Li FT)
  4. TK605 (evolución diaria, scatter, correlaciones)
  5. E2 y E3
  6. Na LM (impacto sobre rendimiento y calidad)
  7. Setpoints recomendados (con rangos min–max)
- **Tabla resumen consolidada** descargable como CSV
- **Meta de rendimiento ajustable** (default 78%)

## Probar localmente (opcional)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abre en http://localhost:8501.

## Deploy en Streamlit Community Cloud (gratis, 5 minutos)

### 1. Crear repositorio en GitHub
1. Ir a https://github.com → "New repository"
2. Nombre: `dashboard-rendimiento-planta`
3. Puede ser **público** o **privado** (Streamlit Cloud soporta ambos en plan gratuito).
4. Subir los archivos `app.py` y `requirements.txt` de esta carpeta.
   - Opción simple: arrastra los archivos al repo desde la web de GitHub.

### 2. Conectar Streamlit Community Cloud
1. Ir a https://share.streamlit.io
2. "Sign in with GitHub"
3. Click "**New app**"
4. Seleccionar:
   - Repository: `<tu-usuario>/dashboard-rendimiento-planta`
   - Branch: `main`
   - Main file path: `app.py`
5. Click "**Deploy**"

En 2–3 minutos la app queda online en una URL del tipo:
```
https://dashboard-rendimiento-planta-<usuario>.streamlit.app
```

### 3. Compartir
- Copia la URL y compártela con el equipo (Gabriel, Alex, Milto, etc.).
- No requiere instalar nada del lado del usuario, solo abrir el link en el navegador.

## Formato de Excel esperado

El archivo debe tener una hoja llamada **`Hoja2`** con las columnas estándar
del análisis (Fecha, Rendimiento AQ, flujos, temperaturas, KPIs por etapa,
washing, V inventario, etc.). El mismo formato del archivo `Datos_planta.xlsx`
que ya usamos.

## Actualización del análisis

Cada vez que el equipo necesite analizar un periodo nuevo:
1. Abrir el dashboard
2. Subir el Excel actualizado
3. Listo — todos los gráficos, correlaciones y setpoints se recalculan automáticamente

## Soporte

- Si el Excel tiene columnas con nombres distintos, se debe ajustar `app.py`.
- Para agregar nuevos análisis, editar `app.py` y hacer push al repo: el deploy
  se actualiza automáticamente en Streamlit Cloud.
