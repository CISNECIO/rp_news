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
        return normalize_source(match.group(1).strip())
    if url and isinstance(url, str):
        try:
            domain = urlparse(url).netloc.replace('www.', '')
            parts = domain.split('.')
            if len(parts) >= 2:
                return normalize_source(parts[0].capitalize())
        except:
            pass
    return "Fuente no identificada"

# ─────────────────────────────────────────────────────────────
# SOURCE NORMALIZATION
# Consolidates variants like "Gestion"/"Gestión", "ThePaypers"/"The paypers",
# "Gan@ Más"/"GanaMás", "Sbs"/"SBS", etc. into canonical names.
# If a raw source isn't in the exact map, heuristic cleanup applies.
# ─────────────────────────────────────────────────────────────
SOURCE_CANONICAL_MAP = {
    # Peruvian press
    'gestion': 'Gestión',
    'gestión': 'Gestión',
    'gestión versión impresa': 'Gestión',
    'gestión - versión impresa': 'Gestión',
    'gestión versión mpresa': 'Gestión',
    'informe de opinión - gestión': 'Gestión',
    'artículo de vodanovic en gestión': 'Gestión',
    'artículo de opinión de ljubica vodanovic en gestión': 'Gestión',
    'gestión y semana económica': 'Gestión',
    'el comercio': 'El Comercio',
    'elcomercio': 'El Comercio',
    'el comerio': 'El Comercio',
    'comercio': 'El Comercio',
    'la república': 'La República',
    'larepublica': 'La República',
    'semanaeconomica': 'Semana Económica',
    'semana económica': 'Semana Económica',
    'perú 21': 'Perú 21',
    'peru21': 'Perú 21',
    'perú retail': 'Perú Retail',
    'perú retail ': 'Perú Retail',
    'perú retail': 'Perú Retail',
    'rpp': 'RPP',
    'andina': 'Andina',
    'infobae': 'Infobae',
    'exitosa': 'Exitosa',
    'exitosanoticias': 'Exitosa',
    'el peruano': 'El Peruano',
    'infomercado': 'Infomercado',
    'diario uno': 'Diario UNO',
    'diariohoy': 'Diario Hoy',
    'expreso': 'Expreso',
    # Fintech-specialized (Peru / LatAm)
    'masfinanzas': 'Más Finanzas',
    'más finanzas': 'Más Finanzas',
    'masfinanz@s': 'Más Finanzas',
    'másfinanz@s': 'Más Finanzas',
    'másfinanzas': 'Más Finanzas',
    'más finanz@s': 'Más Finanzas',
    'más fin@nzas': 'Más Finanzas',
    'másfin@nzas': 'Más Finanzas',
    'másfinan@z': 'Más Finanzas',
    'mas finanzas': 'Más Finanzas',
    'más finanzas': 'Más Finanzas',
    'más': 'Más Finanzas',
    'mas': 'Más Finanzas',
    'microfinanzas': 'Microfinanzas',
    'revistaganamas': 'Gana Más',
    'gan@ más': 'Gana Más',
    'gan@más': 'Gana Más',
    'ganamás': 'Gana Más',
    'ganamas': 'Gana Más',
    'ganam@s': 'Gana Más',
    'gana más': 'Gana Más',
    'gana mas': 'Gana Más',
    'iupana': 'Iupana',
    'latamfintech': 'Latam Fintech Hub',
    'latam fintech': 'Latam Fintech Hub',
    'latam fintech hub': 'Latam Fintech Hub',
    'latam fintech hub': 'Latam Fintech Hub',
    'latamfintech hub': 'Latam Fintech Hub',
    'colombiafintech': 'Colombia Fintech',
    'finnovista': 'Finnovista',
    'fintechnews': 'Fintech News',
    'fintechfutures': 'Fintech Futures',
    'fintech futures': 'Fintech Futures',
    'fintech global': 'Fintech Global',
    'fintech global y finextra': 'Fintech Global',
    'fintech times': 'The Fintech Times',
    'the fintech times': 'The Fintech Times',
    'fintech magazine': 'Fintech Magazine',
    'fintech magnates': 'Fintech Magnates',
    'fintech review': 'Fintech Review',
    'fintechweekly': 'Fintech Weekly',
    'fintech weekly': 'Fintech Weekly',
    # International wires / press
    'reuters': 'Reuters',
    'reuters y cointelegraph': 'Reuters',
    'bloomberg': 'Bloomberg',
    'bloomberglinea': 'Bloomberg Línea',
    'bloomberg línea': 'Bloomberg Línea',
    'bloomberg en línea': 'Bloomberg Línea',
    'financial times': 'Financial Times',
    'ft': 'Financial Times',
    'finextra': 'Finextra',
    'finextra': 'Finextra',
    'fintextra': 'Finextra',
    'paypers': 'The Paypers',
    'the paypers': 'The Paypers',
    'thepaypers': 'The Paypers',
    'the paypers': 'The Paypers',
    'tha paypers': 'The Paypers',
    'pymnts': 'PYMNTS',
    'paymnts': 'PYMNTS',
    'techcrunch': 'TechCrunch',
    'cnbc': 'CNBC',
    'forbes': 'Forbes',
    'forbes perú': 'Forbes',
    'the economist': 'The Economist',
    'the guardian': 'The Guardian',
    'business insider': 'Business Insider',
    'yahoo finance': 'Yahoo Finance',
    'coindesk': 'CoinDesk',
    'cointelegraph': 'CoinTelegraph',
    'the block': 'The Block',
    'americaeconomia': 'América Economía',
    'américa economía': 'América Economía',
    'finanzas y desarrollo del imf': 'IMF',
    # Regulators / institutions
    'sbs': 'SBS',
    'bcrp': 'BCRP',
    'bcrportal': 'BCRP',
    'smv': 'SMV',
    'mef': 'MEF',
    'bis': 'BIS',
    'fca': 'FCA',
    'sec': 'SEC',
    'cepal': 'CEPAL',
    'iadb': 'BID',
    'idbinvest': 'BID Invest',
    'bid': 'BID',
    'asbanc': 'Asbanc',
    'fenacrep': 'Fenacrep',
    'fpcmac': 'FEPCMAC',
    'fepcmac': 'FEPCMAC',
    'bankofengland': 'Bank of England',
    'banco central de canadá': 'Banco de Canadá',
    'centralbanking': 'Central Banking',
    # Corporates / banks
    'bbva': 'BBVA',
    'bbva': 'BBVA',
    'hsbc': 'HSBC',
    'santander': 'Santander',
    'nubank': 'Nubank',
    'nu méxico': 'Nubank',
    'j.p. morgan': 'JP Morgan',
    'jp morgan': 'JP Morgan',
    'bnp paribas': 'BNP Paribas',
    'bnp pariba': 'BNP Paribas',
    'visa': 'Visa',
    'mastercard': 'Mastercard',
    'paypal': 'PayPal',
    'klarna': 'Klarna',
    'revolut': 'Revolut',
    'robinhood': 'Robinhood',
    'etoro': 'eToro',
    'coinbase': 'Coinbase',
    'bitpanda': 'Bitpanda',
    'astropay': 'AstroPay',
    'belvo': 'Belvo',
    'kushki': 'Kushki',
    'ualá': 'Ualá',
    'google': 'Google',
    'google cloud': 'Google',
    # Discard-like / low-signal labels → "Fuente no identificada"
    'es': 'Fuente no identificada',
    'ok': 'Fuente no identificada',
    'blog': 'Fuente no identificada',
    'publishing': 'Fuente no identificada',
    'repe?': 'Fuente no identificada',
    'no va': 'Fuente no identificada',
    'opcional': 'Fuente no identificada',
    'up': 'Fuente no identificada',
    'drive': 'Fuente no identificada',
    'peru': 'Fuente no identificada',
    'gob': 'Fuente no identificada',
    'pl': 'Fuente no identificada',
    'br': 'Fuente no identificada',
    'cloud': 'Fuente no identificada',
    'maximo': 'Fuente no identificada',
    'elias': 'Fuente no identificada',
    'video': 'Fuente no identificada',
    'hola': 'Fuente no identificada',
    'newswriter': 'Fuente no identificada',
    'investor': 'Fuente no identificada',
    'finance': 'Fuente no identificada',
    'city': 'Fuente no identificada',
    'fintech': 'Fuente no identificada',
    'linkedin': 'LinkedIn',
    'facebook': 'Facebook',
    'youtube': 'YouTube',
    'bing': 'Fuente no identificada',
    '24/7': 'Fuente no identificada',
    '6minutos': '6 Minutos',
    'pt50': 'Fuente no identificada',
    'lrt': 'LRT',
    'anna': 'ANNA',
    'ebis': 'EBIS',
    'ey': 'EY',
    'polarg': 'PolArg',
    # Typos / malformed
    'finextra': 'Finextra',
    'www1-folha-uol-com-br': 'Folha',
    'lloys banking group': 'Lloyds Banking Group',
}

def normalize_source(raw):
    """Normalize a source string to a canonical form.
    Applies: trim, strip stray punctuation, exact-map lookup, then heuristic cleanup."""
    if not raw or not isinstance(raw, str):
        return "Fuente no identificada"
    s = raw.strip()
    # Strip trailing/leading stray quotes or punctuation
    s = s.strip('"\'.,; ')
    if not s:
        return "Fuente no identificada"
    # Exact-map lookup (case-insensitive)
    canonical = SOURCE_CANONICAL_MAP.get(s.lower())
    if canonical:
        return canonical
    # Heuristic: Title Case for ALL-CAPS words that look like acronyms (SBS, FCA, BIS)
    if s.isupper() and len(s) <= 4:
        return s  # keep as acronym
    # Heuristic: all-lowercase single word → capitalize
    if s.islower() and ' ' not in s:
        return s.capitalize()
    return s

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
            resumen = str(resumen).strip().replace('\r\n', '\n').replace('\r', '')
            # Collapse 3+ consecutive newlines to 2 (paragraph break)
            import re
            resumen = re.sub(r'\n{3,}', '\n\n', resumen)
            # Short version: single-line for cards
            resumen_short_raw = resumen.replace('\n', ' ')
            resumen_short_raw = re.sub(r'  +', ' ', resumen_short_raw).strip()
            resumen_short = resumen_short_raw[:280] + '...' if len(resumen_short_raw) > 280 else resumen_short_raw
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
    
    print(f" Converted {len(articles)} articles -> {out_path}")
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
