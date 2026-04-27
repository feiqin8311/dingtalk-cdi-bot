import unittest

from app.services.freight_service import calculate_shipping_costs


class FreightServiceTests(unittest.TestCase):
    def test_sea_delivery_uses_weight_basis(self):
        freight, basis = calculate_shipping_costs(
            transportation_method="海派",
            logistics_provider="X",
            country_zh="美国",
            currency="USD",
            total_volume=0.3,
            total_weight=8,
        )
        self.assertEqual(basis, "weight")
        self.assertEqual(freight, 20.0)

    def test_sea_card_uses_volume_basis(self):
        freight, basis = calculate_shipping_costs(
            transportation_method="海卡",
            logistics_provider="X",
            country_zh="美国",
            currency="USD",
            total_volume=1.5,
            total_weight=8,
        )
        self.assertEqual(basis, "volume")
        self.assertEqual(freight, 260.42)


if __name__ == "__main__":
    unittest.main()
