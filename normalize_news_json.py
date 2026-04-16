#!/usr/bin/env python3
"""
One-off normalization pass on data/news.json.

Use this ONLY if you don't want to re-run convert.py from Excel yet.
It re-reads data/news.json, applies source normalization, and writes back.
After a normal convert.py run this becomes a no-op.
"""
import json
import os
import sys

# Reuse the same map / function from convert.py
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from convert import normalize_source

def main():
    base = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base, 'data', 'news.json')
    with open(path, encoding='utf-8') as f:
        data = json.load(f)

    changed = 0
    for a in data['articles']:
        old = a.get('source', '')
        new = normalize_source(old)
        if new != old:
            a['source'] = new
            changed += 1

    data['sources'] = sorted(set(a['source'] for a in data['articles']))

    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"Normalized {changed} article sources.")
    print(f"Unique sources after: {len(data['sources'])}")

if __name__ == '__main__':
    main()
