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
        flask_app.ADMIN_CHAT_ID = "12345"
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

    def test_telegram_purchase_ref_state_machine_and_approval(self):
        user_id = 998877
        user_obj = {'id': user_id, 'first_name': 'TestUser'}

        # 1. Create a Purchase Reference
        res = self.client.post('/api/v1/purchases/create', json={
            'productId': 'freshman_natural_science_y1_full_year',
            'userId': str(user_id)
        })
        self.assertEqual(res.status_code, 200)
        ref = res.get_json()['purchaseReference']

        # 2. User inputs purchase reference in Telegram chat
        handled = flask_app.handle_draft_step(user_obj, ref)
        self.assertTrue(handled)

        draft = flask_app._get_draft(user_id)
        self.assertEqual(draft['step'], 'ref_confirm')
        self.assertEqual(draft['purchaseReference'], ref)

        # 3. User clicks Confirm
        callback_update = {
            'callback_query': {
                'id': 'cb_123',
                'data': f'confirm_ref:{ref}',
                'from': user_obj,
                'message': {'chat': {'id': user_id}, 'message_id': 1}
            }
        }
        cb_handled = flask_app.handle_callback(callback_update)
        self.assertTrue(cb_handled)

        draft_after_confirm = flask_app._get_draft(user_id)
        self.assertEqual(draft_after_confirm['step'], 'phone')

    def test_repeat_purchases_and_language_switching(self):
        user_id = 112233
        user_obj = {'id': user_id, 'first_name': 'RepeatUser'}

        # Select Afaan Oromoo language
        cb_lang = {
            'callback_query': {
                'id': 'cb_lang_1',
                'data': 'lang:om',
                'from': user_obj,
                'message': {'chat': {'id': user_id}, 'message_id': 10}
            }
        }
        self.assertTrue(flask_app.handle_callback(cb_lang))
        self.assertEqual(flask_app.get_user_lang(user_id), 'om')

        # Verify start message in Oromiffa
        msg = flask_app.get_start_message(user_id)
        self.assertIn("Application Barnoota Temhiro", msg)

        # First purchase submit
        draft = {'package': 'freshman', 'phone': {'number': '251911111111', 'verified': True}, 'method': 'cbe', 'name': 'User 1', 'link': 'https://mbreciept.cbe.com.et/test', 'txid': 'FT123456', 'step': 'done'}
        req1 = flask_app.submit_request(user_id, draft)
        self.assertIsNotNone(req1)

        # /start again for a second purchase
        start_update = {'message': {'chat': {'id': user_id}, 'from': user_obj, 'text': '/start'}}
        self.assertTrue(flask_app.handle_message(start_update))

        # Second purchase submit
        draft2 = {'package': 'freshman', 'phone': {'number': '251911111111', 'verified': True}, 'method': 'telebirr', 'name': 'User 1', 'link': 'https://transactioninfo.ethiotelecom.et/receipt/test', 'txid': 'DG123456', 'step': 'done'}
        req2 = flask_app.submit_request(user_id, draft2)
        self.assertIsNotNone(req2)
        self.assertNotEqual(req1['requestId'], req2['requestId'])

    def test_admin_guided_addpackage_flow(self):
        admin_id = 12345
        admin_obj = {'id': admin_id, 'first_name': 'Admin'}

        # Step 1: Send /addpackage without args
        up1 = {'message': {'chat': {'id': admin_id}, 'from': admin_obj, 'text': '/addpackage'}}
        self.assertTrue(flask_app.handle_message(up1))
        draft = flask_app._get_draft(admin_id)
        self.assertEqual(draft.get('admin_step'), 'addpkg_name')

        # Step 2: Send Label/Name
        up2 = {'message': {'chat': {'id': admin_id}, 'from': admin_obj, 'text': 'Freshman Natural Science - Semester 1'}}
        self.assertTrue(flask_app.handle_message(up2))
        draft = flask_app._get_draft(admin_id)
        self.assertEqual(draft.get('admin_step'), 'addpkg_key')
        self.assertEqual(draft.get('addpkg_label'), 'Freshman Natural Science - Semester 1')

        # Step 3: Send Tag/Key
        up3 = {'message': {'chat': {'id': admin_id}, 'from': admin_obj, 'text': 'custom_freshman_sem1'}}
        self.assertTrue(flask_app.handle_message(up3))
        draft = flask_app._get_draft(admin_id)
        self.assertEqual(draft.get('admin_step'), 'addpkg_price')
        self.assertEqual(draft.get('addpkg_key'), 'custom_freshman_sem1')

        # Step 4: Send Price Tag in ETB
        up4 = {'message': {'chat': {'id': admin_id}, 'from': admin_obj, 'text': '300'}}
        self.assertTrue(flask_app.handle_message(up4))
        
        # Verify package saved
        pkg = flask_app.get_package_by_key('custom_freshman_sem1')
        self.assertIsNotNone(pkg)
        self.assertEqual(pkg['label'], 'Freshman Natural Science - Semester 1')
        self.assertEqual(pkg['priceCents'], 30000)

if __name__ == '__main__':
    unittest.main()




