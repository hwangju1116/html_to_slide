import unittest
from unittest.mock import MagicMock, patch
import os
from app.slide_digitizer import digitize_slide_image

class TestSlideDigitizer(unittest.TestCase):
    @patch("app.slide_digitizer.get_genai_client")
    def test_digitize_slide_image_mock(self, mock_client_fn):
        mock_client = MagicMock()
        mock_client_fn.return_value = mock_client

        mock_resp_html = MagicMock()
        mock_resp_html.text = "```html\n<div class=\"slide\"><h1 contenteditable=\"true\">Test Title</h1></div>\n```"
        
        mock_resp_pptx = MagicMock()
        mock_resp_pptx.text = '```python\nfrom pptx import Presentation\nprs = Presentation()\nslide = prs.slides.add_slide(prs.slide_layouts[6])\nprs.save(r"/tmp/mock_digitized.pptx")\n```'

        mock_client.models.generate_content.side_effect = [mock_resp_html, mock_resp_pptx]

        with patch("app.slide_fidelity_critic.audit_slide_fidelity") as mock_audit:
            from app.schemas import SlideFidelityReport
            mock_audit.return_value = SlideFidelityReport(passed=True, score=9, zero_omission_verified=True)
            res = digitize_slide_image(
                b"fake_png_bytes",
                output_basename="mock_digitized",
                workspace_dir="/tmp",
            )
        self.assertTrue(res["success"])
        self.assertIn("contenteditable", res["html_content"])
        self.assertTrue(os.path.exists("/tmp/mock_digitized.html"))

if __name__ == "__main__":
    unittest.main()
