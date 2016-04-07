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
from pycel.excellib import mid
from pycel.excellib import year

class Test_Year(unittest.TestCase):
    def setup(self):
        pass

    def test_year_must_be_integer(self):
        with self.assertRaises(TypeError):
            year('2016', 1, 1)

    def test_month_must_be_integer(self):
        with self.assertRaises(TypeError):
            year(2016, '1', 1)

    def test_day_must_be_integer(self):
        with self.assertRaises(TypeError):
            year(2016, 1, '1')

    def test_year_must_be_positive(self):
        with self.assertRaises(ValueError):
            year(-1, 1, 1)

    def test_year_must_have_less_than_10000(self):
        with self.assertRaises(ValueError):
            year(10000, 1, 1)

    def test_result_must_be_positive(self):
        with self.assertRaises(ArithmeticError):
            year(1900, 1, -1)

    def test_not_stricly_positive_month_substracts(self):
        self.assertEqual(year(2009, -1, 1), year(2008, 11, 1))

    def test_not_stricly_positive_day_substracts(self):
        self.assertEqual(year(2009, 1, -1), year(2008, 12, 30))

    def test_month_superior_to_12_change_year(self):
        self.assertEqual(year(2009, 14, 1), year(2010, 2, 1))

    def test_day_superior_to_365_change_year(self):
        self.assertEqual(year(2009, 1, 400), year(2010, 2, 4))

    def test_year_between_1900_and_9999(self):
        self.assertEqual(year(2008, 114, 3), 42889)

class Test_Mid(unittest.TestCase):
    def setUp(self):
        pass

    def test_start_num_must_be_integer(self):
        with self.assertRaises(TypeError):
            mid('Romain', 1.1, 2)

    def test_num_chars_must_be_integer(self):
        with self.assertRaises(TypeError):
            mid('Romain', 1, 2.1)

    def test_start_num_must_be_superior_or_equal_to_1(self):
        with self.assertRaises(ValueError):
            mid('Romain', 0, 3)

    def test_num_chars_must_be_positive(self):
        with self.assertRaises(ValueError):
            mid('Romain', 1, -1)

    def test_mid(self):
        self.assertEqual(mid('Romain', 2, 9), 'main')
        

class Test_Round(unittest.TestCase):
    def setUp(self):
        pass

    def test_nb_must_be_number(self):
        with self.assertRaises(TypeError):
            round('er', 1)

    def test_nb_digits_must_be_number(self):
        with self.assertRaises(TypeError):
            round(2.323, 'ze')

    def test_positive_number_of_digits(self):
        self.assertEqual(xround(2.675, 2), 2.68)

    def test_negaive_number_of_digits(self):
        self.assertEqual(xround(2352.67, -2), 2400) 

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
