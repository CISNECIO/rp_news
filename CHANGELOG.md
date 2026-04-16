# CHANGELOG — SAFR / RP News

## ⚠️ Actualización importante (bugfix)

La versión anterior tenía un bug serio: los `onerror` de imagen inyectaban HTML mal escapado, lo que causaba placeholders rotos (iconito 🖼️ + texto `">` suelto + badges "INVERSIONES" flotando fuera de lugar). **Esta versión lo arregla**.

**Causa raíz:** `onerror="this.outerHTML='<div class=\"ph ...\">...'"` — las comillas dobles del `class="..."` dentro del placeholder cortaban prematuramente el atributo `onerror`.

**Fix:** nueva función `phFallback(img, articleId, extraClass)` que busca el artículo por id y hace `replaceChild` limpio. Aplicado en las 6 ubicaciones (hero-lead, hero-mid, cat-venture-lead, cat-headline-thumb, feed-card, modal).

**Verificado:**
- ✅ 0 patrones `this.outerHTML` residuales
- ✅ 6 `phFallback` llamadas
- ✅ Div balance 203/203
- ✅ JS syntax válido (node --check)
- ✅ Python syntax válido (py_compile)
- ✅ Test de integración con jsdom: 3 placeholders se insertan correctamente, texto limpio, sin garbage
- ✅ Incluye guarda anti-loop (`imgEl._phFailed`) por si el fallback a su vez fallara

---

## Archivos modificados / nuevos

| Archivo                     | Acción    | Notas                                                                                          |
| --------------------------- | --------- | ---------------------------------------------------------------------------------------------- |
| `index.html`                | Modificado | Todos los cambios de UI/UX (CSS + JS embebido). 3838 líneas. Div balance verificado.            |
| `convert.py`                | Modificado | Nuevo `SOURCE_CANONICAL_MAP` + `normalize_source()`. `extract_source()` llama a normalize.     |
| `normalize_news_json.py`    | Nuevo      | One-off: aplica `normalize_source` al `news.json` existente (útil antes de correr Excel→JSON). |
| `data/news.json`            | Regenerado | 1,525 artículos normalizados en fuentes. Contenido de artículos sin cambios.                   |

## Cómo desplegar

Opción A (rápida, sin Excel):
1. Copiar los 4 archivos a tu carpeta del proyecto respetando la subcarpeta `data/`.
2. Abrir `index.html` (o `python -m http.server 8000` y ir a `http://localhost:8000`).

Opción B (correr desde Excel):
1. Copiar `index.html`, `convert.py`, `normalize_news_json.py` al proyecto.
2. `%run convert.py` en Spyder (usa el nuevo mapa automáticamente).
3. `normalize_news_json.py` queda como utilidad standalone si lo vuelves a necesitar.

---

## Cambios aprobados implementados

### 1. Indicador "última actualización" visible
- **Dónde**: header, junto al botón Focus.
- **Qué**: pill con punto verde pulsante cuando es reciente (<24h), dorado estático cuando es stale.
- **Texto relativo**: "Actualizado hace 23 min" / "Actualizado ayer" / "Actualizado hace 3 días".
- **Tooltip**: muestra el timestamp completo al pasar el cursor.

### 2. Legibilidad en listas densas
- `.cat-text-item`, `.hero-headline-item`, `.feed-card`, `.cat-headline-item` → más line-height, más padding interno, gap mayor.
- Metas ahora usan `--gray-500` (oscurecido) en vez de `--gray-400`.

### 3. Stats con insight temporal
Arriba de la vista Estadísticas, strip de 4 cards:
- **Volumen · últ. 30d**: cantidad + Δ% vs los 30 días anteriores (flecha roja/verde).
- **Promedio diario**: artículos/día en el periodo.
- **Categoría en alza**: la categoría con mayor crecimiento relativo (mínimo 3 artículos para evitar ruido). Clickeable → filtra.
- **Día pico**: fecha con más artículos en el periodo. Clickeable → filtra.

### 4. Stats cross-linking a artículos filtrados
- **Timeline clickeable**: cada barra (días/semanas/meses) filtra por ese rango exacto. Tooltip lo anuncia.
- **Insight cards clickeables**: categoría en alza y día pico filtran directo.
- **Cross-links ya existentes** (región, categorías, temas, fuentes) siguen funcionando.

### 5. Contraste en texto secundario
- `--gray-400`: `#9aa4b4` → `#7b8699`
- `--gray-500`: `#6b7789` → `#5a6779`

### 6. Manejo de imágenes vacías
Sistema `.ph` (placeholder con identidad):
- Placeholder coloreado por categoría: Fintech=azul, Banca=navy, Pagos=dorado, Regulación=verde, Inversiones=slate.
- Iniciales de la fuente como letra grande (p.ej. "GE" para Gestión, "RE" para Reuters).
- Tag de categoría en esquina inferior izquierda.
- Detección de URLs genéricas (patrones `no-image`, `placeholder`, Finextra default, etc.) → trata como vacío.
- Aplicado en hero-lead, hero-mid, cat-venture-lead, cat-headline-thumb, feed-card, modal.
- **phFallback()** maneja errores de carga limpiamente (reemplazo por `replaceChild`, no por HTML injection).

### 7. Hover states más fuertes
Todas las cards que abren artículo ahora tienen barra lateral acento al hover:
- `.feed-card`, `.cat-text-item`, `.hero-headline-item`, `.cat-headline-item` → `border-left: 2px solid` acento + desplazamiento +2px + background sutil.

### 8. "Limpiar filtros" más visible
- `.filter-status__clear` y `.filter-chips__clear` convertidos a pill rojo con ícono, hover con lift y sombra.

### 9. Saved Views (vistas guardadas)
Nuevo botón "★ Vistas ▾" en header, junto al chip de actualización.
- Captura: búsqueda, categoría, fuente, región, preset de fechas, rango explícito.
- Persistencia: `localStorage.bcrp_saved_views_v1`, máximo 8 vistas.
- Panel dropdown con lista + "Guardar vista actual".
- Click → aplica; X → elimina.
- Prompt para nombrar (con default inteligente).
- Toast de confirmación al guardar.
- Click-outside cierra el panel.

### 10. Normalización de fuentes
- Hecho upstream en `convert.py` (decisión técnica: JSON limpio = stats correctos).
- Mapa de ~200 variantes → canónicas. Casos consolidados:
  - `Gestion` + `Gestión` + `Gestión versión impresa` + ... → **Gestión** (719 artículos)
  - `Revistaganamas` + `Gan@ Más` + `GanaMás` + ... → **Gana Más** (272)
  - `The Paypers` + `Thepaypers` + `The paypers` + `Paypers` → **The Paypers** (174)
  - `Bloomberglinea` + `Bloomberg en Línea` → **Bloomberg Línea**
  - `Sbs` / `SBS` → **SBS**
  - `Bcrp` / `BCRP` / `Bcrportal` → **BCRP**
- Heurística fallback: mayúsculas consistentes para siglas (SBS, BIS, FCA); capitalización para palabras sueltas.
- Resultado: de 200+ fuentes únicas a **163** (1,525 artículos normalizados).

---

## Features nuevas tangenciales

- **Chip de rango de fechas** en filter chips (📅 15 Mar — 21 Mar) cuando haces click en una barra de la timeline o en el insight "Día pico". Removible.
- **Toast** genérico reutilizable (`showToast(msg)`).

---

## No implementado (respetando tus restricciones)

- ❌ Daily briefing
- ❌ Más filtros visibles
- ❌ Filtro manual de rango de fechas (el chip aparece solo cuando lo dispara algo más)
- ❌ Cambios al layout editorial del hero
- ❌ Contador de artículos por sección en home
- ❌ Sort-by adicional
- ❌ "Alta relevancia" manual

---

## Notas de mantenimiento

### Al agregar una fuente nueva
Si aparece una variante no mapeada, agrégala a `SOURCE_CANONICAL_MAP` en `convert.py`:
```python
'cinco dias': 'Cinco Días',
```

### Al resetear saved views (DevTools Console)
```js
localStorage.removeItem('bcrp_saved_views_v1');
```

### Al ajustar el periodo del insight strip
En `buildInsightStrip()`, variable `WINDOW = 30` (días).

### Claves localStorage usadas
- `bcrp_last_visit` — banner "X noticias nuevas" (ya existía)
- `bcrp_saved_views_v1` — vistas guardadas (nueva)

### Session storage
- `focusMode` — estado del modo Focus entre recargas (ya existía)
