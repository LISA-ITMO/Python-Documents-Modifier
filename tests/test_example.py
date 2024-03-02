import unittest

from src.CodecoveTest import CodecoveTest


class SimpleTestCase(unittest.TestCase):
    # TODO: Тест для проверки работы, удалить
    def test_one(self):
        self.assertEqual(23, CodecoveTest.calc())