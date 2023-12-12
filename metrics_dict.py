import pprint
import pandas as pd

# in review, always cut by metric name first


def metric_groupby_dimension(metrics, dimension):
    """reviewing a metric by a dimension"""
    by_dimension = {}
    if isinstance(metrics, str):
        by_dimension.setdefault(metrics, [])
        for item in dimension:
            by_dimension[metrics].append(item)
    elif isinstance(metrics, list):
        for metric in metrics:
            by_dimension.setdefault(metric, [])
            for item in dimension:
                by_dimension[metric].append(item)
    else:
        raise TypeError('metrics must be a string or a list')
    return by_dimension


gms = ['threep_net_ord_gms', 'fba_net_ord_gms', 'mfn_net_ord_gms']
units_sold = ['threep_net_ord_units', 'fba_net_ord_units', 'mfn_net_ord_units']
ads = ['ads_spend', 'ads_attributed_ops', 'sp_spend', 'sp_attributed_ops',
       'sb_spend', 'sb_attributed_ops', 'sd_spend', 'sd_attributed_ops']



metric_mapping = {
        'threep_net_ord_gms': 'Net Ordered GMS',
        'fba_net_ord_gms': 'FBA Net Ordered GMS',
        'mfn_net_ord_gms': 'MFN Net Ordered GMS',
        'threep_net_ord_units': 'Net Ordered Units',
        'fba_net_ord_units': 'FBA Net Ordered Units',
        'mfn_net_ord_units': 'MFN Net Ordered Units'
}

seller_origin_mapping = {
            'KR': ['KR', 'KP'],
            'TW': ['TW'],
            'VN': ['VN'],
            'RSEA': ['SG', 'TH', 'ID', 'MY', 'PH', 'KH'],
            'CN': ['CN', 'HK', 'MO'],
            'IN': ['IN'],
            'LATAM': ['BR', 'MX']
        }

arcs = {
        'Established': {
            'AGG': ['US', 'CA', 'UK', 'DE', 'FR', 'IT', 'ES', 'JP'],
            'NA': ['US', 'CA'],
            'EU': ['UK', 'DE', 'FR', 'IT', 'ES'],
            'JP': ['JP']
            },
        'WW': {
            'AGG': ['US', 'CA', 'UK', 'DE', 'FR', 'IT', 'ES', 'JP', 'IN', 'BR',
                    'MX', 'PL', 'SE', 'AU', 'SG', 'NL', 'AE', 'SA', 'BE', 'TR',
                    'EG'],
            'NA': ['US', 'CA'],
            'EU': ['UK', 'DE', 'FR', 'IT', 'ES'],
            'JP': ['JP'],
            'Emerging': ['IN', 'BR', 'MX', 'PL', 'SE', 'AU', 'SG', 'NL', 'AE',
                         'SA', 'BE', 'TR', 'EG']
            }
        }

seller_age = {
        'New': {
            'DSR': ['DSR'],
            'SSR': ['SSR-A', 'SSR-P', 'SSR-Inv']
            },
        'Existing': {
            'Managed': ['SAS-CORE', 'TBAM', 'SAM', 'MASS'],
            'Non-Managed': ['Non-Managed']
            }
        }


## Wrtie to a Json template
#file = open('metric.py', 'w')
#file.write('ww_group_by_region = ' + pprint.pformat(group_by_region['WW']) + '\n')
#file.write('est_ww_group_by_region = ' + pprint.pformat(group_by_region['Established']) + '\n')
#file.close()
