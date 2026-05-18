# RP News

Visualizador estático de noticias financieras para BCRP / SAFR.

El sitio no usa build step ni servidor permanente: `index.html` contiene el frontend y lee los datos desde `data/news.json`.

## Flujo Principal

1. Reemplazar o actualizar el Excel local `Sample.xlsx`.
2. Ejecutar `python convert.py`.
3. Abrir `index.html` o servir la carpeta con `python -m http.server 8000`.

`convert.py` normaliza el Excel y genera `data/news.json`, que es el único JSON de noticias usado por la web. Los Excel `Sample*.xlsx` son insumos locales y están ignorados por Git.

## Estructura

```text
rp-news/
├── index.html              # Aplicación web estática
├── convert.py              # Excel -> data/news.json
├── data/
│   ├── news.json           # Datos activos del sitio
│   └── market.json         # Indicadores de mercado, opcional
├── functions/
│   └── _middleware.js      # Protección Cloudflare Pages
├── update_rp_news.py       # Automatización de conversión + commit + push
├── run_rp_news_daily.bat   # Wrapper para Programador de tareas de Windows
└── update_market.py        # Generador opcional de data/market.json
```

Archivo local esperado para conversión:

```text
Sample.xlsx                 # Excel fuente local, no versionado
```

## Requisitos

- Python 3.10+ recomendado
- Paquetes: `pandas`, `openpyxl`
- Navegador moderno

Instalación mínima:

```bash
pip install pandas openpyxl
```

Para indicadores de mercado:

```bash
pip install yfinance
```

## Ejecutar Localmente

```bash
python convert.py
python -m http.server 8000
```

Abrir `http://localhost:8000`.

## Automatización

`update_rp_news.py` hace:

1. `git pull --ff-only origin main`
2. `python convert.py`
3. `git add data/news.json`
4. Commit y push solo si cambió el JSON generado

Los logs se escriben en `logs/task.log`, pero `logs/` está ignorado por Git.

En Windows, programar `run_rp_news_daily.bat`; el `.bat` usa `py -3` si existe y si no usa `python`.

## GitHub y Credenciales

No guardes credenciales en este repositorio. Usa una de estas opciones:

- Git Credential Manager: ejecutar `git pull` o `git push` y completar el login del navegador.
- GitHub CLI: `gh auth login`, luego usar Git normalmente.
- Token personal: usarlo solo en el prompt seguro de Git, nunca dentro de archivos del proyecto.

## Datos

Cada artículo en `data/news.json` tiene:

- `id`
- `title`
- `source`
- `date`
- `date_sort`
- `url`
- `summary`
- `summary_short`
- `image_url`
- `category`
- `keywords`
- `region`
