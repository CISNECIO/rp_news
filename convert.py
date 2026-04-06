#!/usr/bin/env python3
"""
RP News - Excel to JSON Converter
Reads the news Excel file and produces a normalized JSON for the frontend.
Usage: python convert.py [path_to_excel]
Default: looks for Sample.xlsx in the same directory.
"""

import pandas as pd
import json
import sys
import re
import os
from urllib.parse import urlparse
from datetime import datetime

def extract_source(title, url):
    """Extract source from title parentheses, then URL domain, then fallback."""
    match = re.search(r'\(([^)]+)\)\s*$', title or '')
    if match:
        return match.group(1).strip()
    if url and isinstance(url, str):
        try:
            domain = urlparse(url).netloc.replace('www.', '')
            parts = domain.split('.')
            if len(parts) >= 2:
                return parts[0].capitalize()
        except:
            pass
    return "Fuente no identificada"

def clean_title(title):
    """Remove source parenthetical from end of title and strip whitespace."""
    if not title:
        return "Sin título"
    title = title.strip().replace('\n', ' ').replace('\r', '')
    title = re.sub(r'\s*\([^)]+\)\s*$', '', title)
    return title.strip()

def infer_category(title, keywords, existing):
    """Use classification if exists, else infer from keywords/title.
    All categories normalized to: Fintech, Banca, Pagos, Regulación, Inversiones."""
    ALLOWED = ['Fintech', 'Banca', 'Pagos', 'Regulación', 'Inversiones']
    
    if existing and isinstance(existing, str) and existing.strip():
        cat = existing.strip()
        mapping = {
            'Banking': 'Banca',
            'banking': 'Banca',
            'Banca': 'Banca',
            'Pagos': 'Pagos',
            'Pagos ': 'Pagos',
            'Préstamos': 'Banca',
            'Prestamos': 'Banca',
            'Regulación': 'Regulación',
            'Regulacion': 'Regulación',
            'Inversiones': 'Inversiones',
            'Fintech': 'Fintech',
            'Criptoactivos': 'Fintech',
            'IA / tecnología': 'Fintech',
            'IA / Tecnología': 'Fintech',
            'Inclusión financiera': 'Banca',
            'Inclusión Financiera': 'Banca',
        }
        result = mapping.get(cat, None)
        if result:
            return result
        # Try partial match
        cat_lower = cat.lower()
        if 'regul' in cat_lower: return 'Regulación'
        if 'pago' in cat_lower: return 'Pagos'
        if 'banc' in cat_lower or 'bank' in cat_lower or 'préstam' in cat_lower or 'prestam' in cat_lower: return 'Banca'
        if 'invers' in cat_lower or 'bolsa' in cat_lower: return 'Inversiones'
        if 'fintech' in cat_lower or 'cripto' in cat_lower or 'tecnol' in cat_lower or 'ia' in cat_lower: return 'Fintech'
    
    text = f"{title} {keywords}".lower()
    if any(w in text for w in ['regulación', 'sbs', 'indecopi', 'normativa', 'ley', 'regulador']):
        return 'Regulación'
    if any(w in text for w in ['pago', 'pagos', 'transferencia', 'pasarela']):
        return 'Pagos'
    if any(w in text for w in ['préstamo', 'crédito', 'bnpl', 'deuda', 'financiamiento', 'banco', 'bancario', 'banca', 'banking']):
        return 'Banca'
    if any(w in text for w in ['inversión', 'bolsa', 'acciones', 'valores']):
        return 'Inversiones'
    if any(w in text for w in ['fintech', 'startup', 'neobank', 'cripto', 'bitcoin', 'blockchain', 'inteligencia artificial', 'machine learning', 'tecnología']):
        return 'Fintech'
    return 'Banca'  # fallback to most common

def parse_keywords(kw_str):
    """Parse keywords string into a clean list."""
    if not kw_str or not isinstance(kw_str, str):
        return []
    keywords = [k.strip().rstrip('.') for k in kw_str.split(',')]
    return [k for k in keywords if k and len(k) > 1]

def convert(excel_path):
    df = pd.read_excel(excel_path)
    articles = []

    # Handle both old and new column names
    title_col = 'Título del artículo' if 'Título del artículo' in df.columns else 'Titulo'
    keywords_col = 'Palabras claves' if 'Palabras claves' in df.columns else 'Palabras clave'
    region_col = 'Region' if 'Region' in df.columns else 'Región'
    
    for i, row in df.iterrows():
        fecha = row.get('Fecha')
        if pd.isna(fecha):
            fecha_str = "Fecha no disponible"
            fecha_sort = "1900-01-01"
        else:
            if isinstance(fecha, datetime):
                fecha_str = fecha.strftime('%d %b %Y')
                fecha_sort = fecha.strftime('%Y-%m-%d')
            else:
                fecha_str = str(fecha)
                fecha_sort = str(fecha)
        
        raw_title = str(row.get(title_col, '')).strip()
        source = extract_source(raw_title, row.get('URL'))
        title = clean_title(raw_title)
        
        url = row.get('URL')
        url = str(url).strip() if pd.notna(url) else None
        
        resumen = row.get('Resumen')
        if pd.notna(resumen):
            resumen = str(resumen).strip().replace('\n', ' ').replace('\r', '')
            resumen_short = resumen[:280] + '...' if len(resumen) > 280 else resumen
        else:
            resumen = None
            resumen_short = None
        
        img_url = row.get('URL imagen')
        img_url = str(img_url).strip() if pd.notna(img_url) else None
        
        keywords_raw = row.get(keywords_col)
        keywords = parse_keywords(str(keywords_raw) if pd.notna(keywords_raw) else '')
        
        category = infer_category(
            raw_title,
            str(keywords_raw) if pd.notna(keywords_raw) else '',
            row.get('Clasificación')
        )

        # Region: Local / Internacional
        region_raw = row.get(region_col)
        if pd.notna(region_raw) and str(region_raw).strip():
            region = str(region_raw).strip()
        else:
            region = 'Local'  # default
        
        articles.append({
            'id': i,
            'title': title,
            'source': source,
            'date': fecha_str,
            'date_sort': fecha_sort,
            'url': url,
            'summary': resumen,
            'summary_short': resumen_short,
            'image_url': img_url,
            'category': category,
            'keywords': keywords,
            'region': region,
        })
    
    # Sort by date descending
    articles.sort(key=lambda x: x['date_sort'], reverse=True)
    
    # Extract unique categories and sources
    categories = sorted(set(a['category'] for a in articles))
    sources = sorted(set(a['source'] for a in articles))
    regions = sorted(set(a['region'] for a in articles))
    
    output = {
        'generated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'total': len(articles),
        'categories': categories,
        'sources': sources,
        'regions': regions,
        'articles': articles,
    }
    
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'news.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    
    print(f"✓ Converted {len(articles)} articles → {out_path}")
    print(f"  Categories: {', '.join(categories)}")
    print(f"  Sources: {', '.join(sources)}")
    print(f"  Regions: {', '.join(regions)}")
    return output

if __name__ == '__main__':
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Sample.xlsx')
    if not os.path.exists(path):
        print(f"Error: File not found: {path}")
        sys.exit(1)
    convert(path)
