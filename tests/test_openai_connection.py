import os
import sys
from pathlib import Path
import unittest

from openai import OpenAI

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config.settings import BASE_URL_AI, OPENAI_API_KEY, OPENAI_MODEL


class TestOpenAIConnection(unittest.TestCase):
    def test_openai_chat_connection(self):
        if not OPENAI_API_KEY:
            self.skipTest("OPENAI_API_KEY is not set")

        client_kwargs = {"api_key": OPENAI_API_KEY}
        if BASE_URL_AI:
            client_kwargs["base_url"] = BASE_URL_AI

        client = OpenAI(**client_kwargs)

        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[
                {
                    "role": "system",
                    "content": "Reply with a short confirmation only.",
                },
                {
                    "role": "user",
                    "content": "Reply with exactly: connection ok",
                }
            ],
            max_completion_tokens=10,
            temperature=0,
        )

        content = response.choices[0].message.content or ""
        self.assertTrue(content.strip())
        self.assertIn("connection", content.lower())


if __name__ == "__main__":
    unittest.main()
