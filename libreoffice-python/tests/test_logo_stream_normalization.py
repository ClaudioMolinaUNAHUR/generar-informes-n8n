import base64
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from utils.helpers import normalize_logo_stream


class LogoStreamNormalizationTests(unittest.TestCase):
    def test_dict_payload_with_base64_is_normalized_to_bytesio(self):
        payload = {"logo_base64": base64.b64encode(b"fake-image-data").decode("ascii")}

        stream = normalize_logo_stream(payload)

        self.assertIsNotNone(stream)
        self.assertTrue(hasattr(stream, "seek"))
        self.assertTrue(hasattr(stream, "read"))


if __name__ == "__main__":
    unittest.main()
