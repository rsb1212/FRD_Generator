import unittest
import xml.etree.ElementTree as ET
from app import app, ExtractorFactory, SOPExtractor, RGTExtractor, clean_and_dedupe_frd
from app import generate_drawio_xml


class FRDAppTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True, SECRET_KEY='test-secret')
        self.client = app.test_client()

    def test_root_redirects_to_login_when_not_authenticated(self):
        response = self.client.get('/', follow_redirects=False)
        self.assertEqual(response.status_code, 302)
        self.assertIn('/login', response.headers['Location'])

    def test_login_get_and_post(self):
        res_get = self.client.get('/login')
        self.assertEqual(res_get.status_code, 200)

        res_invalid = self.client.post('/login', data={'username': 'wrong', 'password': 'bad'}, follow_redirects=True)
        self.assertIn(b'Invalid username or password', res_invalid.data)

        res_valid = self.client.post('/login', data={'username': 'admin', 'password': 'admin123'}, follow_redirects=False)
        self.assertEqual(res_valid.status_code, 302)
        self.assertEqual(res_valid.headers['Location'], '/')

    def test_logout(self):
        with self.client as c:
            with c.session_transaction() as sess:
                sess['logged_in'] = True
            response = c.get('/logout', follow_redirects=True)
            self.assertIn(b'signed out', response.data)

    def test_ai_status_endpoint(self):
        with self.client as c:
            with c.session_transaction() as sess:
                sess['logged_in'] = True
            response = c.get('/ai-status')
            self.assertEqual(response.status_code, 200)
            json_data = response.get_json()
            self.assertIn('local_nlp', json_data)

    def test_generate_requires_rgt_file(self):
        with self.client as c:
            with c.session_transaction() as sess:
                sess['logged_in'] = True
            response = c.post('/generate', data={}, follow_redirects=True)
            self.assertIn(b'Please upload the RGT file.', response.data)

    def test_ai_compare_requires_frd_first(self):
        with self.client as c:
            with c.session_transaction() as sess:
                sess['logged_in'] = True
            app.config['LAST_SOP_TEXT'] = ""
            app.config['LAST_RGT_TEXT'] = ""
            response = c.get('/ai-compare', follow_redirects=True)
            self.assertIn(b'Generate an FRD first', response.data)

    def test_download_requires_frd_first(self):
        with self.client as c:
            with c.session_transaction() as sess:
                sess['logged_in'] = True
            app.config['LAST_FRD'] = None
            response = c.get('/download', follow_redirects=True)
            self.assertIn(b'Generate an FRD first', response.data)

    def test_extractor_factory_and_strategies(self):
        sop = ExtractorFactory.create_sop_extractor(["PURPOSE: Test Purpose", "SCOPE: Test Scope"])
        self.assertIsInstance(sop, SOPExtractor)
        self.assertEqual(sop.get_purpose(), "Test Purpose")

        # sheets → rows → cells: one sheet with a header row then a data row
        rgt = ExtractorFactory.create_rgt_extractor([
            [["Functionality", "Expected", "Requirement"], ["Feature A", "Expected A", "Req A"]]
        ])
        self.assertIsInstance(rgt, RGTExtractor)
        self.assertEqual(len(rgt.get_delta_features()), 1)

    def test_clean_and_dedupe_frd(self):
        raw_frd = {
            'stakeholders': ['PS', 'PS', 'CCM', 'ccm'],
            'process_steps': [
                {'actor': 'Customer', 'description': 'Step 1'},
                {'actor': 'Customer', 'description': 'Step 1'}
            ]
        }
        cleaned = clean_and_dedupe_frd(raw_frd)
        self.assertEqual(len(cleaned['stakeholders']), 2)
        self.assertEqual(len(cleaned['process_steps']), 1)
        self.assertIn('mermaid_diagram', cleaned)
        self.assertIn('process_flow_chain', cleaned)

    def test_drawio_xml_uses_vertical_swimlanes_and_required_shapes(self):
        xml = generate_drawio_xml([
            {'actor': 'Customer', 'description': 'Start request'},
            {'actor': 'System', 'description': 'Validate request'},
            {'actor': 'System', 'description': 'Save record'},
            {'actor': 'Customer', 'description': 'Generate report and complete'},
        ])
        cells = ET.fromstring(xml).findall('.//mxCell')
        lanes = [cell for cell in cells if cell.get('style') == 'swimlane;whiteSpace=wrap;html=1;']

        self.assertEqual(len(lanes), 2)
        self.assertTrue(any(cell.get('style') == 'ellipse;whiteSpace=wrap;html=1;fillColor=#dae8fc;' for cell in cells))
        self.assertTrue(any(cell.get('style') == 'rhombus;whiteSpace=wrap;html=1;fillColor=#ffe6cc;' for cell in cells))
        self.assertTrue(any('shape=cylinder3;' in cell.get('style', '') for cell in cells))
        self.assertTrue(any('shape=document;' in cell.get('style', '') for cell in cells))


if __name__ == '__main__':
    unittest.main()
