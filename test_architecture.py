import unittest
import os
import shutil
import tempfile
import json
import flask_app

class TestTemhiroBotArchitecture(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        flask_app.DATA_DIR = self.temp_dir
        flask_app.USERS_FILE = os.path.join(self.temp_dir, "users.json")
        flask_app.VOUCHERS_FILE = os.path.join(self.temp_dir, "vouchers.json")
        flask_app.ENTITLEMENTS_FILE = os.path.join(self.temp_dir, "entitlements.json")
        flask_app.PURCHASES_FILE = os.path.join(self.temp_dir, "purchases.json")
        flask_app.PACKAGES_FILE = os.path.join(self.temp_dir, "packages.json")
        flask_app.app.testing = True
        self.client = flask_app.app.test_client()

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_canonical_products_pricing_and_entitlements(self):
        sem_product = flask_app.CANONICAL_PRODUCTS.get('freshman_natural_science_y1_sem1')
        self.assertIsNotNone(sem_product)
        self.assertEqual(sem_product['priceCents'], 30000)
        self.assertEqual(sem_product['entitlements'], ['freshman_natural_science_y1_sem1'])

        full_year_product = flask_app.CANONICAL_PRODUCTS.get('freshman_natural_science_y1_full_year')
        self.assertIsNotNone(full_year_product)
        self.assertEqual(full_year_product['priceCents'], 50000)
        self.assertEqual(sorted(full_year_product['entitlements']), ['freshman_natural_science_y1_sem1', 'freshman_natural_science_y1_sem2'])

        coc_product = flask_app.CANONICAL_PRODUCTS.get('coc_medical')
        self.assertIsNotNone(coc_product)
        self.assertEqual(coc_product['priceCents'], 40000)

    def test_purchase_reference_api(self):
        res = self.client.post('/api/v1/purchases/create', json={
            'productId': 'freshman_natural_science_y1_full_year'
        })
        self.assertEqual(res.status_code, 200)
        data = res.get_json()
        self.assertTrue(data.get('success'))
        ref = data.get('purchaseReference')
        self.assertTrue(ref.startswith('TH-'))
        purchase = data.get('purchase')
        self.assertEqual(purchase.get('priceCents'), 50000)

        get_res = self.client.get(f'/api/v1/purchases/{ref}')
        self.assertEqual(get_res.status_code, 200)
        get_data = get_res.get_json()
        self.assertTrue(get_data.get('success'))
        get_purchase = get_data.get('purchase')
        self.assertEqual(get_purchase['productId'], 'freshman_natural_science_y1_full_year')
        self.assertEqual(get_purchase['priceCents'], 50000)
        self.assertEqual(sorted(get_purchase['entitlements']), ['freshman_natural_science_y1_sem1', 'freshman_natural_science_y1_sem2'])

if __name__ == '__main__':
    unittest.main()


