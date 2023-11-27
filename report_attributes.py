date_dimension = ['activity_week', 'activity_month', 'activity_quarter', 'activity_year']

time_prefix = ['wtd', 'mtd', 'qtd', 'ytd']

region_dimension = ['region', 'marketplace_id', 'marketplace_name']

seller_origin_dimension = ['seller_origin', 'seller_origin_level1', 'seller_origin_level2']

new_seller_dimension = ['seller_age', 'launch_channel', 'launch_sub_channel', 'is_seller_fraud']

existing_seller_dimension = ['seller_age', 'esm_team']

gms_perf_metrics = [
    'threep_net_ord_gms', 'fba_net_ord_gms', 'mfn_net_ord_gms',
    'threep_net_ord_units', 'fba_net_ord_units', 'mfn_net_ord_units'
]

ads_perf_metrics = [
    'sp_spend', 'sp_attributed_gms', 'sp_clicks', 'sp_impressions'
    'sb_spend', 'sb_attributed_gms', 'sb_clicks', 'sb_impressions'
    'sd_spend', 'sd_attributed_gms', 'sd_clicks', 'sd_impressions'
]

promotion_perf_metrics = [
    'promotion_ops', 'deal_ops', 'deal_bd_ops', 'deal_dotd_ops', 'deal_ld_ops',
    'promotion_units', 'deal_units', 'deal_bd_units', 'deal_dotd_units', 'deal_ld_units'
]