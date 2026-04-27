def calculate_shipping_costs(transportation_method: str, logistics_provider: str, country_zh: str,
                             currency: str, total_volume: float, total_weight: float) -> tuple[float, str]:
    exchange_rate = 7.2 if currency == "USD" else 1.0
    multiplier_num1 = 1250
    multiplier_num2 = 1
    freight_basis = "volume"

    if "海卡" in transportation_method:
        sea_card_rates = {
            "加拿大": 800,
            "日本": 800,
            "德国": 732 if logistics_provider == "北京世纪卓越快递服务有限公司（龙舟）" else 1200,
            "英国": 634 if logistics_provider == "北京世纪卓越快递服务有限公司（龙舟）" else 1200,
        }
        multiplier_num1 = sea_card_rates.get(country_zh, 1250)
        multiplier_num2 = max(1.0, total_volume)
        freight_basis = "volume"
    elif "海派" in transportation_method:
        sea_delivery_rates = {
            "美国": 12,
            "加拿大": 12,
            "澳洲": 10,
            "日本": 10,
            "英国": 11,
        }
        multiplier_num1 = sea_delivery_rates.get(country_zh, 12)
        multiplier_num2 = max(12.0, total_weight)
        freight_basis = "weight"
    elif "空派" in transportation_method:
        air_delivery_rates = {
            "美国": 45,
            "德国": 42,
        }
        multiplier_num1 = air_delivery_rates.get(country_zh, 45)
        multiplier_num2 = max(12.0, total_weight)
        freight_basis = "weight"
    elif "卡航" in transportation_method:
        truck_delivery_rates = {
            "英国": 15.5,
            "德国": 16,
        }
        multiplier_num1 = truck_delivery_rates.get(country_zh, 15.5)
        multiplier_num2 = total_weight
        freight_basis = "weight"
    elif "美国海运" in transportation_method:
        sea_transport_rates = {
            "美国": 500,
        }
        multiplier_num1 = sea_transport_rates.get(country_zh, 1250)
        multiplier_num2 = max(1.0, total_volume)
        freight_basis = "volume"

    return round(multiplier_num1 * multiplier_num2 / exchange_rate, 2), freight_basis
