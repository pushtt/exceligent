import pprint
import pandas as pd

metrics = {
        'threep_net_ord_gms': 'Net Ordered GMS',
        'fba_net_ord_gms': 'FBA Net Ordered GMS',
        'mfn_net_ord_gms': 'MFN Net Ordered GMS',
        'net_ord_units': 'Net Ordered Units',
        'fba_net_ord_units': 'FBA Net Ordered Units',
        'mfn_net_ord_units': 'MFN Net Ordered Units'
}


seller_origins = {
            'KR': ['KR'],
            'TW': ['TW'],
            'VN': ['VN'],
            'RSEA': [
                'SG', 'TH', 'ID', 'MY', 'PH', 'KH'
                ],
            'CN': ['CN'],
            'IN': ['IN'],
            'LATAM': [
                'BR', 'MX'
                ]
        }

arcs = {
        'Established': {
            'NA': ['US', 'CA'],
            'EU': ['UK', 'DE', 'FR', 'IT', 'ES'],
            'JP': ['JP']
            },
        'WW': {
            'NA': ['US', 'CA'],
            'EU': ['UK', 'DE', 'FR', 'IT', 'ES'],
            'JP': 'JP',
            'Emerging': []
            }
        }

seller_age = {
        'New': {
            'DSR': 'DSR',
            'SSR': ['SSR-A', 'SSR-P']
            },
        'Existing': {
            'Managed': ['SAS-CORE', 'TBAM', 'SAM', 'MASS'],
            'Non-Managed': ['Non-Managed']
            }
        }


group_by_region = {}
group_by_region_by_country = {}

for arc, region in arcs.items():
    group_by_region.setdefault(arc, {})
    for metric in metrics:
        group_by_region[arc].setdefault(metric, [])
        group_by_region[arc][metric] += region.keys()

# Wrtie to a Json template
file = open('metric.py', 'w')
file.write('ww_group_by_region = ' + pprint.pformat(group_by_region['WW']) + '\n')
file.write('est_ww_group_by_region = ' + pprint.pformat(group_by_region['Established']) + '\n')
file.close()
