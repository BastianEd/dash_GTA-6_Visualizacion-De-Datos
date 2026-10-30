# 🌴 GTA VI · Dashboard para Rockstar

Dashboard interactivo de visualización de datos que analiza los comentarios públicos de tres videos de **GTA VI** para responder una pregunta central:

> **¿Qué tema debería priorizar Rockstar al comunicar GTA VI?**

Proyecto de la asignatura **ADY1104 · Visualización de Datos**.

🔗 **Demo en vivo:** [dash-gta-6-visualizacion-de-datos.onrender.com](https://dash-gta-6-visualizacion-de-datos.onrender.com/)

> ⏳ La app está en el plan gratuito de Render: si estuvo inactiva, la primera carga puede tardar unos segundos.

---

## 🎯 Hallazgo principal

> *Rockstar puede comunicar con confianza la historia y los personajes, pero primero debe aclarar las dudas sobre plataformas y rendimiento.*

---

## 🖥️ Vistas del dashboard

| Vista | Qué muestra |
|---|---|
| **1 · Resumen ejecutivo** | Indicadores clave, comparación central de sentimiento por tema (barras divergentes), decisión recomendada para Rockstar, cómo comprobarla, nivel de "Se busca" por tema y citas reales de la audiencia. |
| **2 · Temas y comparaciones** | Cinco preguntas con su conclusión: de qué habla la gente, qué temas generan más críticas y elogios, si cambia según el video, dónde se concentran las críticas (tabla tema × video) y sobre qué temas se hacen preguntas. |
| **3 · Evidencia y calidad** | Tabla de comentarios reales (ordenable y paginada), búsqueda por palabra, filtro de preguntas y una sección sobre los límites y la confiabilidad de las etiquetas. |

### Funcionalidades

- **Panel de control** con filtros por video, sentimiento y tema principal.
- **Búsqueda de texto** sin distinguir mayúsculas ni tildes (`grafico` encuentra `gráfico`).
- **Filtro de tipo de comentario:** todos, solo preguntas o sin preguntas.
- **Indicadores de confiabilidad:** se marca con ⚠ cualquier grupo con menos de 30 comentarios, y las celdas del mapa tema × video con menos de 10 comentarios no se pintan.
- **Cuenta regresiva** al lanzamiento oficial (19/11/2026).
- **Estilo visual Vice City:** paleta neón sobre noche violeta, ilustración SVG con sol retro y palmeras, tipografías Anton, Space Grotesk y Source Sans 3.

---

## 🗂️ Estructura del repositorio

```
.
├── app.py                          # Aplicación Dash (carga de datos, gráficos, vistas y callbacks)
├── data/
│   └── GTA_VI_Datos_Visualizacion.xlsx   # Dataset de comentarios (hoja "Datos")
├── infografia_gta_vi_exacta.pdf    # Infografía de una página
├── render.yaml                     # Blueprint de despliegue en Render
├── requirements.txt                # Dependencias de Python
└── README.md
```

---

## 🛠️ Tecnologías

- [Python](https://www.python.org/)
- [Dash](https://dash.plotly.com/) y [Plotly](https://plotly.com/python/) para la interfaz y los gráficos
- [pandas](https://pandas.pydata.org/) y [openpyxl](https://openpyxl.readthedocs.io/) para leer y procesar el Excel
- [Gunicorn](https://gunicorn.org/) (producción en Render) y [Waitress](https://docs.pylonsproject.org/projects/waitress/) (servidor WSGI en Windows)

---

## 🚀 Ejecución local

### 1. Clonar el repositorio

```bash
git clone https://github.com/BastianEd/dash_GTA-6_Visualizacion-De-Datos.git
cd dash_GTA-6_Visualizacion-De-Datos
```

### 2. Crear un entorno virtual (recomendado) e instalar dependencias

```bash
python -m venv venv

# Windows
venv\Scripts\activate
# macOS / Linux
source venv/bin/activate

pip install -r requirements.txt
```

### 3. Verificar el dataset

El archivo `GTA_VI_Datos_Visualizacion.xlsx` debe estar en la carpeta `data/`.

### 4. Ejecutar

```bash
python app.py
```

Abre [http://127.0.0.1:8088](http://127.0.0.1:8088) en tu navegador.

**Alternativa en Windows (servidor WSGI):**

```bash
waitress-serve --listen=127.0.0.1:8088 app:server
```

---

## ☁️ Despliegue en Render

El proyecto incluye un `render.yaml` (Blueprint), por lo que se puede desplegar así:

1. Sube el repositorio a GitHub.
2. En [Render](https://render.com/), crea un nuevo **Blueprint** y conecta el repositorio.
3. Render detectará `render.yaml` y desplegará la aplicación automáticamente.

El objeto `server` expuesto en `app.py` es el que consume el servidor WSGI (`app:server`).

---

## 📊 Sobre los datos

- **Muestra:** comentarios públicos y anonimizados de **3 videos** de GTA VI.
- **Etiquetas:** el sentimiento (Positivo, Neutral, Negativo) y el tema principal fueron asignados por un programa con reglas simples; **no tienen validación humana**.
- **Preguntas:** se detectan comentarios que contienen una pregunta (`es_pregunta`).

### Limitaciones

- Son reacciones a tres videos, no la experiencia de jugar GTA VI.
- Es una muestra: no representa a toda la comunidad ni demuestra causas.
- Cada comentario tiene un único tema, aunque pueda hablar de varios.
- Con grupos pequeños, un solo comentario mueve mucho el porcentaje.

---

## 🎓 Contexto académico

| | |
|---|---|
| **Asignatura** | ADY1104 · Visualización de Datos |
| **Entregables** | Dashboard web (este repositorio) e infografía en PDF |
| **Autor** | Nicolas Fonseca - Cristobal segovia - Bastian Rubio |

---

## 📄 Aviso

Proyecto académico con fines educativos. GTA VI y Rockstar Games son marcas de sus respectivos propietarios; este proyecto no está afiliado ni respaldado por ellos.

