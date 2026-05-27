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
import math
import unicodedata
from collections import Counter, defaultdict
from urllib.parse import urlparse
from datetime import datetime, timedelta

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

def normalize_external_url(value):
    """Return a clean absolute URL, or None for flags/placeholders such as Si/X."""
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace('&amp;', '&')
    if text.lower() in {'si', 'sí', 's\u00ed', 'x', 'no', 'nan', 'none', 'null'}:
        return None
    if text.startswith('www.'):
        text = 'https://' + text
    parsed = urlparse(text)
    if parsed.scheme not in {'http', 'https'} or not parsed.netloc:
        return None
    return text

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


STOPWORDS = {
    'a', 'al', 'ante', 'antes', 'asi', 'bajo', 'cada', 'como', 'con', 'contra',
    'cuando', 'de', 'del', 'desde', 'donde', 'durante', 'e', 'el', 'ella',
    'ellas', 'ellos', 'en', 'entre', 'era', 'es', 'esa', 'ese', 'eso', 'esta',
    'estan', 'estar', 'este', 'esto', 'estos', 'fue', 'ha', 'han', 'hasta',
    'hay', 'la', 'las', 'le', 'les', 'lo', 'los', 'mas', 'mientras', 'muy',
    'no', 'o', 'para', 'pero', 'por', 'porque', 'que', 'se', 'segun', 'ser',
    'si', 'sin', 'sobre', 'su', 'sus', 'tambien', 'tras', 'un', 'una', 'uno',
    'y', 'the', 'and', 'for', 'from', 'into', 'with', 'that', 'this', 'will',
    'financiero', 'financiera', 'financieros', 'financieras', 'sistema',
    'banco', 'bancos', 'empresa', 'empresas', 'mercado', 'sector', 'nuevo',
    'nueva', 'nuevos', 'nuevas', 'peru', 'peruano', 'peruana', 'articulo',
    'noticia', 'noticias', 'fintech', 'banca', 'pagos', 'regulacion',
    'inversiones'
}

SOURCE_TIERS = {
    'official': {
        'SBS', 'BCRP', 'BIS', 'SEC', 'FCA', 'SMV', 'MEF', 'Banco de Canadá',
        'Bank of England', 'CEPAL', 'BID', 'BID Invest'
    },
    'market_intel': {
        'Bloomberg', 'Bloomberg Línea', 'Reuters', 'Financial Times',
        'The Economist', 'Finextra', 'The Paypers', 'PYMNTS',
        'The Fintech Times', 'Fintech Futures'
    },
    'specialized': {
        'Iupana', 'Latam Fintech Hub', 'Colombia Fintech', 'Finnovista',
        'Más Finanzas', 'Gana Más', 'Semana Económica', 'Gestión',
        'Microfinanzas'
    },
}

ENTITY_GROUPS = {
    'regulators': [
        ('SBS', ['sbs', 'superintendencia de banca']),
        ('BCRP', ['bcrp', 'banco central de reserva']),
        ('BIS', ['bis', 'banco de pagos internacionales']),
        ('SEC', ['sec', 'securities and exchange commission']),
        ('FCA', ['fca', 'financial conduct authority']),
        ('SMV', ['smv', 'superintendencia del mercado de valores']),
        ('MEF', ['mef', 'ministerio de economia']),
        ('Banco de Inglaterra', ['bank of england', 'banco de inglaterra']),
        ('Banco de Canadá', ['bank of canada', 'banco de canada']),
    ],
    'companies': [
        ('Nubank', ['nubank', 'nu mexico', 'nu méxico']),
        ('Visa', ['visa']),
        ('Mastercard', ['mastercard']),
        ('PayPal', ['paypal']),
        ('BBVA', ['bbva']),
        ('Coinbase', ['coinbase']),
        ('Klarna', ['klarna']),
        ('Revolut', ['revolut']),
        ('Ualá', ['uala', 'ualá']),
        ('Belvo', ['belvo']),
        ('Kushki', ['kushki']),
        ('AstroPay', ['astropay']),
        ('eToro', ['etoro']),
        ('Google', ['google']),
        ('Lloyds Banking Group', ['lloyds banking group', 'lloyds']),
    ],
    'countries': [
        ('Perú', ['peru', 'perú']),
        ('Brasil', ['brasil', 'brazil']),
        ('México', ['mexico', 'méxico']),
        ('Colombia', ['colombia']),
        ('Chile', ['chile']),
        ('Argentina', ['argentina']),
        ('Estados Unidos', ['estados unidos', 'united states', 'eeuu', 'usa']),
        ('Reino Unido', ['reino unido', 'united kingdom', 'uk']),
        ('Unión Europea', ['union europea', 'unión europea', 'european union', 'ue']),
        ('Canadá', ['canada', 'canadá']),
        ('China', ['china']),
    ],
}

RISK_RULES = [
    ('Regulación', ['regulacion', 'normativa', 'ley', 'supervision', 'supervisor', 'licencia', 'sbs', 'smv', 'sec', 'fca', 'bis']),
    ('Ciberseguridad / fraude', ['fraude', 'ciberseguridad', 'ciberataque', 'phishing', 'estafa', 'mulas', 'lavado', 'aml', 'riesgo operacional']),
    ('Criptoactivos', ['cripto', 'crypto', 'bitcoin', 'stablecoin', 'stablecoins', 'tokenizacion', 'tokenización', 'activos digitales', 'blockchain']),
    ('IA / modelos', ['inteligencia artificial', ' ia ', 'machine learning', 'modelo', 'algoritmo', 'automatizacion', 'automatización']),
    ('Pagos e infraestructura', ['pagos', 'pago', 'interoperabilidad', 'transferencia', 'instantaneo', 'instantáneo', 'wallet', 'billetera', 'qr']),
    ('Open finance', ['open finance', 'open banking', 'finanzas abiertas', 'banca abierta', 'apis', 'api']),
    ('Crédito / deuda', ['credito', 'crédito', 'prestamo', 'préstamo', 'bnpl', 'deuda', 'morosidad', 'financiamiento']),
    ('Protección al consumidor', ['consumidor', 'usuario', 'reclamo', 'transparencia', 'comisiones', 'indecopi', 'proteccion', 'protección']),
    ('Estabilidad financiera', ['estabilidad financiera', 'liquidez', 'solvencia', 'sistemico', 'sistémico', 'capital', 'riesgo financiero']),
    ('Mercados / inversión', ['inversion', 'inversión', 'valores', 'bolsa', 'acciones', 'trading', 'asset management']),
]


def strip_accents(text):
    return ''.join(
        ch for ch in unicodedata.normalize('NFD', str(text or ''))
        if unicodedata.category(ch) != 'Mn'
    )


def norm_text(text):
    return strip_accents(text).lower()


def tokenize(text):
    words = re.findall(r'[a-záéíóúñü0-9]{3,}', str(text or '').lower(), flags=re.I)
    clean = []
    for word in words:
        key = norm_text(word)
        if key not in STOPWORDS and not key.isdigit():
            clean.append(key)
    return clean


def parse_date_sort(date_sort):
    try:
        return datetime.strptime(str(date_sort), '%Y-%m-%d')
    except Exception:
        return None


def text_contains_alias(text, alias):
    alias_norm = norm_text(alias).strip()
    if not alias_norm:
        return False
    if re.fullmatch(r'[a-z0-9]{2,4}', alias_norm):
        return re.search(r'(^|[^a-z0-9])' + re.escape(alias_norm) + r'([^a-z0-9]|$)', text) is not None
    return alias_norm in text


def detect_entities(article):
    text = norm_text(
        f"{article.get('title', '')} {article.get('summary', '')} "
        f"{' '.join(article.get('keywords', []))} {article.get('source', '')}"
    )
    found = {}
    for group, entities in ENTITY_GROUPS.items():
        hits = []
        for name, aliases in entities:
            if any(text_contains_alias(text, alias) for alias in aliases):
                hits.append(name)
        found[group] = hits
    return found


def detect_risk_tags(article):
    text = norm_text(
        f"{article.get('title', '')} {article.get('summary', '')} "
        f"{' '.join(article.get('keywords', []))} {article.get('category', '')}"
    )
    tags = []
    for label, aliases in RISK_RULES:
        if any(text_contains_alias(text, alias) for alias in aliases):
            tags.append(label)
    return tags


def source_tier(source):
    if source in SOURCE_TIERS['official']:
        return 'official'
    if source in SOURCE_TIERS['market_intel']:
        return 'market_intel'
    if source in SOURCE_TIERS['specialized']:
        return 'specialized'
    if source == 'Fuente no identificada':
        return 'unknown'
    return 'general'


def source_score(source):
    return {
        'official': 100,
        'market_intel': 82,
        'specialized': 72,
        'general': 52,
        'unknown': 25,
    }[source_tier(source)]


def article_terms(article, max_terms=10):
    terms = []
    seen = set()
    for kw in article.get('keywords', []):
        clean = norm_text(kw).strip()
        if clean and clean not in STOPWORDS and len(clean) > 2 and clean not in seen:
            seen.add(clean)
            terms.append(kw.strip())
    text_terms = tokenize(f"{article.get('title', '')} {article.get('summary_short', '')}")
    for token in text_terms:
        if token not in seen:
            seen.add(token)
            terms.append(token)
        if len(terms) >= max_terms:
            break
    return terms[:max_terms]


def slugify_value(text):
    value = norm_text(text)
    value = re.sub(r'[^a-z0-9]+', '-', value).strip('-')
    return value or 'sin-cluster'


def label_from_terms(terms, fallback):
    useful = [t.strip() for t in terms if norm_text(t).strip() not in STOPWORDS]
    if not useful:
        return fallback or 'Sin cluster'
    label = ' · '.join(useful[:2])
    return label[:80]


def build_story_key(article):
    tokens = tokenize(article.get('title', ''))
    if len(tokens) < 3:
        tokens = tokenize(f"{article.get('title', '')} {' '.join(article.get('keywords', []))}")
    return ' '.join(sorted(set(tokens[:8]))[:6]) or slugify_value(article.get('title', ''))


def enrich_articles(articles):
    """Add transparent scoring, entity, risk, novelty, and cluster fields."""
    latest = max((parse_date_sort(a.get('date_sort')) for a in articles), default=None)
    latest = latest if latest else datetime.now()
    current_start = latest - timedelta(days=29)
    previous_start = latest - timedelta(days=59)
    previous_end = current_start - timedelta(days=1)

    article_term_lists = {}
    doc_freq = Counter()
    current_terms = Counter()
    previous_terms = Counter()

    for a in articles:
        terms = article_terms(a)
        article_term_lists[a['id']] = terms
        unique_norm_terms = {norm_text(t) for t in terms if norm_text(t)}
        doc_freq.update(unique_norm_terms)
        dt = parse_date_sort(a.get('date_sort'))
        if dt:
            if current_start <= dt <= latest:
                current_terms.update(unique_norm_terms)
            elif previous_start <= dt <= previous_end:
                previous_terms.update(unique_norm_terms)

    momentum_by_term = {}
    for term, cur in current_terms.items():
        prev = previous_terms.get(term, 0)
        pct = ((cur - prev) / prev * 100) if prev else (100 if cur else 0)
        volume_bonus = min(40, cur * 4)
        momentum_by_term[term] = max(0, min(100, 45 + (pct / 6) + volume_bonus))

    story_counts = Counter(build_story_key(a) for a in articles)
    cluster_counts = Counter()
    cluster_labels = {}
    for a in articles:
        terms = article_term_lists[a['id']]
        risk_tags = detect_risk_tags(a)
        if risk_tags:
            cluster_label = risk_tags[0]
        else:
            cluster_label = label_from_terms(terms, a.get('category'))
        cluster_id = slugify_value(cluster_label)
        a['_cluster_id_tmp'] = cluster_id
        a['_cluster_label_tmp'] = cluster_label
        cluster_counts[cluster_id] += 1
        cluster_labels[cluster_id] = cluster_label

    for a in articles:
        terms = article_term_lists[a['id']]
        entities = detect_entities(a)
        entity_total = sum(len(v) for v in entities.values())
        risk_tags = detect_risk_tags(a)
        risk_score = min(100, len(risk_tags) * 22)
        strategic_score = min(100, 20 + risk_score + entity_total * 8)
        if a.get('region') == 'Local':
            strategic_score = min(100, strategic_score + 8)
        if a.get('category') == 'Regulación':
            strategic_score = min(100, strategic_score + 12)

        story_size = story_counts[build_story_key(a)]
        novelty = max(25, 100 - min(15, story_size - 1) * 6)
        if story_size == 1:
            novelty = 100

        term_momentum = [
            momentum_by_term.get(norm_text(t), 45)
            for t in terms
            if norm_text(t)
        ]
        momentum = round(sum(term_momentum) / len(term_momentum)) if term_momentum else 45

        src_score = source_score(a.get('source'))
        entity_score = min(100, entity_total * 18)
        region_score = 70 if a.get('region') == 'Local' else 55
        components = {
            'strategic_relevance': round(strategic_score),
            'source_quality': round(src_score),
            'novelty': round(novelty),
            'momentum': round(momentum),
            'risk_pressure': round(risk_score),
            'entity_density': round(entity_score),
            'regional_relevance': round(region_score),
        }
        signal = round(
            components['strategic_relevance'] * 0.24 +
            components['source_quality'] * 0.15 +
            components['novelty'] * 0.14 +
            components['momentum'] * 0.18 +
            components['risk_pressure'] * 0.16 +
            components['entity_density'] * 0.08 +
            components['regional_relevance'] * 0.05
        )
        signal = max(0, min(100, signal))

        reasons = []
        if risk_tags:
            reasons.append('Riesgo: ' + ', '.join(risk_tags[:2]))
        if entity_total:
            flat_entities = [e for group in entities.values() for e in group]
            reasons.append('Entidades: ' + ', '.join(flat_entities[:3]))
        tier = source_tier(a.get('source'))
        if tier in ('official', 'market_intel'):
            reasons.append('Fuente de alta señal')
        if momentum >= 70:
            reasons.append('Tema en aceleración')
        if novelty >= 90:
            reasons.append('Historia poco repetida')

        cluster_id = a.pop('_cluster_id_tmp')
        cluster_label = a.pop('_cluster_label_tmp')
        a.update({
            'topic_terms': terms,
            'entities': entities,
            'entity_count': entity_total,
            'risk_tags': risk_tags,
            'source_tier': tier,
            'novelty_score': round(novelty),
            'momentum_score': round(momentum),
            'story_cluster_size': story_size,
            'topic_cluster': {
                'id': cluster_id,
                'label': cluster_label,
                'size': cluster_counts[cluster_id],
            },
            'score_components': components,
            'signal_score': signal,
            'signal_level': 'Alta' if signal >= 75 else ('Media' if signal >= 50 else 'Baja'),
            'signal_reasons': reasons[:4],
        })

    top_clusters = [
        {
            'id': cid,
            'label': cluster_labels[cid],
            'count': count,
        }
        for cid, count in cluster_counts.most_common(20)
    ]
    return {
        'scoring_version': 'daifr-signal-v1',
        'latest_date': latest.strftime('%Y-%m-%d'),
        'current_window_days': 30,
        'top_clusters': top_clusters,
        'risk_tags': [label for label, _ in RISK_RULES],
    }

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
        
        img_url = normalize_external_url(row.get('URL imagen'))
        
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
    intelligence = enrich_articles(articles)
    
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
        'intelligence': intelligence,
        'articles': articles,
    }
    
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'news.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    
    print(f" Converted {len(articles)} articles -> {out_path}")
    print(f"  Categories: {', '.join(categories)}")
    print(f"  Sources: {', '.join(sources)}")
    print(f"  Regions: {', '.join(regions)}")
    print(f"  Intelligence: {intelligence['scoring_version']} / {len(intelligence['top_clusters'])} clusters")
    return output

if __name__ == '__main__':
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Sample.xlsx')
    if not os.path.exists(path):
        print(f"Error: File not found: {path}")
        sys.exit(1)
    convert(path)
