import json
from pathlib import Path
import pandas as pd

ANODE_XLSX = Path('/Users/syahnaz.omar/Library/CloudStorage/OneDrive-Personal/Desktop/sesco/Working File/Anode Trending.xlsx')
SKG_XLSX = Path('/Users/syahnaz.omar/Library/CloudStorage/OneDrive-Personal/Desktop/sesco/Working File/SKG Database/SKG Asset Dimension 2025.xlsx')

OUT_ANODE = Path('data/anode_trending')
OUT_SKG = Path('data/skg')
OUT_ANODE.mkdir(parents=True, exist_ok=True)
OUT_SKG.mkdir(parents=True, exist_ok=True)

# Keep raw Inputs layout to preserve parser behavior.
inputs_raw = pd.read_excel(ANODE_XLSX, sheet_name='Inputs', header=None)
inputs_raw.to_csv(OUT_ANODE / 'inputs_raw.csv', index=False, header=False)

m1 = pd.read_excel(ANODE_XLSX, sheet_name='M1PQ-A')
m1.to_csv(OUT_ANODE / 'm1pq_a.csv', index=False)

remaining = pd.read_excel(ANODE_XLSX, sheet_name='Remaining Life')
remaining.to_csv(OUT_ANODE / 'remaining_life.csv', index=False)

# Optional JSON mirrors for API/static usage.
(OUT_ANODE / 'm1pq_a.json').write_text(m1.to_json(orient='records', force_ascii=False), encoding='utf-8')
(OUT_ANODE / 'remaining_life.json').write_text(remaining.to_json(orient='records', force_ascii=False), encoding='utf-8')

psc = pd.read_excel(SKG_XLSX, sheet_name='PSC')
psc.to_csv(OUT_SKG / 'psc.csv', index=False)
(OUT_SKG / 'psc.json').write_text(psc.to_json(orient='records', force_ascii=False), encoding='utf-8')

manifest = {
    'anode': ['inputs_raw.csv', 'm1pq_a.csv', 'remaining_life.csv', 'm1pq_a.json', 'remaining_life.json'],
    'skg': ['psc.csv', 'psc.json'],
}
(Path('data') / 'manifest.json').write_text(json.dumps(manifest, indent=2), encoding='utf-8')

print('Export complete')
