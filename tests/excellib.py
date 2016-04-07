import os
import sys
import unittest


dir = os.path.dirname(__file__)
path = os.path.join(dir, '../src')
sys.path.insert(0, path)

from pycel.excellib import match
from pycel.excellib import mod
from pycel.excellib import count
from pycel.excellib import xround

class Test_Round(unittest.TestCase):
    def setUp(self):
        pass

    def test_nb_must_be_number(self):
        with self.assertRaises(TypeError):
            round('er', 1)

    def test_nb_digits_must_be_number(self):
        with self.assertRaises(TypeError):
            round(2.323, 'ze')

    def test_round_output(self):
        self.assertEqual(xround(2.675, 2), 2.68) 

class Test_Count(unittest.TestCase):
    def setUp(self):
        pass

    def test_without_nested_booleans(self):
        self.assertEqual(count([1, 2, 'e'], True, 'r'), 3)

    def test_with_nested_booleans(self):
        self.assertEqual(count([1, True, 'e'], True, 'r'), 2)

    def test_with_text_representations(self):
        self.assertEqual(count([1, '2.2', 'e'], True, '20'), 4)

class Test_Mod(unittest.TestCase):
    def setUp(self):
        pass

    def test_argument_validity(self):
        with self.assertRaises(TypeError):
            mod(2.2, 1)
        with self.assertRaises(TypeError):
            mod(2, 1.1)

    def test_output_value(self):
        self.assertEqual(mod(10, 4), 2)

class Test_Match(unittest.TestCase):
    def setUp(self):
        pass

    def test_ascending_numeric(self):
        # Closest inferior value is found
        self.assertEqual(match(5, [1, 3.3, 5]), 3)

        # Not ascending arrays raise exception
        with self.assertRaises(Exception):
            match(3, [10, 9.1, 6.23, 1])
        with self.assertRaises(Exception):
            match(3, [10, 3.3, 5, 2])

    def test_exact_numeric(self):
        # Value is found
        self.assertEqual(match(5, [10, 3.3, 5.0], 0), 3)

        # Value not found raises Exception
        with self.assertRaises(ValueError):
            match(3, [10, 3.3, 5, 2], 0)

    def test_descending_numeric(self):
        # Closest superior value is found
        self.assertEqual(match(8, [10, 9.1, 6.2], -1), 2)

        # Non descending arrays raise exception
        with self.assertRaises(Exception):
            match(3, [1, 3.3, 5, 6], -1)
        with self.assertRaises(Exception):
            match(3, [10, 3.3, 5, 2], -1)

    def test_ascending_string(self):
        # Closest inferior value is found
        self.assertEqual(match('rars', ['a', 'AAB', 'rars']), 3)

        # Not ascending arrays raise exception
        with self.assertRaises(Exception):
            match(3, ['rars', 'aab', 'a'])
        with self.assertRaises(Exception):
            match(3, ['aab', 'a', 'rars'])

    def test_exact_string(self):
        # Value is found
        self.assertEqual(match('a', ['aab', 'a', 'rars'], 0), 2)

        # Value not found raises Exception
        with self.assertRaises(ValueError):
            match('b', ['aab', 'a', 'rars'], 0)

    def test_descending_string(self):
        # Closest superior value is found
        self.assertEqual(match('a', ['c', 'b', 'a'], -1), 3)

        # Non descending arrays raise exception
        with self.assertRaises(Exception):
            match('a', ['a', 'aab', 'rars'], -1)
        with self.assertRaises(Exception):
            match('a', ['aab', 'a', 'rars'], -1)

    def test_ascending_boolean(self):
        # Closest inferior value is found
        self.assertEqual(match(True, [False, False, True]), 3)

        # Not ascending arrays raise exception
        with self.assertRaises(Exception):
            match(False, [True, False, False])
        with self.assertRaises(Exception):
            match(True, [False, True, False])

    def test_exact_boolean(self):
        # Value is found
        self.assertEqual(match(False, [True, False, False], 0), 2)

        # Value not found raises Exception
        with self.assertRaises(ValueError):
            match(False, [True, True, True], 0)

    def test_descending_boolean(self):
        # Closest superior value is found
        self.assertEqual(match(False, [True, False, False], -1), 3)

        # Non descending arrays raise exception
        with self.assertRaises(Exception):
            match(False, [False, False, True], -1)
        with self.assertRaises(Exception):
            match(True, [False, True, False], -1)
 
if __name__ == '__main__':
    unittest.main()
