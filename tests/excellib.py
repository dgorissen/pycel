import os
import sys
import unittest


dir = os.path.dirname(__file__)
path = os.path.join(dir, '../src')
sys.path.insert(0, path)

from pycel.excellib import match
from pycel.excellib import mod
from pycel.excellib import count
from pycel.excellib import countif
from pycel.excellib import countifs
from pycel.excellib import xround
from pycel.excellib import mid
from pycel.excellib import date
from pycel.excellib import yearfrac
from pycel.excellib import isNa


class Test_IsNa(unittest.TestCase):
    # This function might need more solid testing
    def setup(self):
        pass

    def test_isNa_false(self):
        self.assertFalse(isNa('2 + 1'))

    def test_isNa_true(self):
        self.assertTrue(isNa('x + 1'))


class Test_Yearfrac(unittest.TestCase):
    def setup(self):
        pass

    def test_start_date_must_be_number(self):
        with self.assertRaises(TypeError):
            yearfrac('not a number', 1)

    def test_end_date_must_be_number(self):
        with self.assertRaises(TypeError):
            yearfrac(1, 'not a number')

    def test_start_date_must_be_positive(self):
        with self.assertRaises(ValueError):
            yearfrac(-1, 0)

    def test_end_date_must_be_positive(self):
        with self.assertRaises(ValueError):
            yearfrac(0, -1)

    def test_basis_must_be_between_0_and_4(self):
        with self.assertRaises(ValueError):
            yearfrac(1, 2, 5)

    def test_yearfrac_basis_0(self):
        self.assertAlmostEqual(yearfrac(date(2008, 1, 1), date(2015, 4, 20)), 7.30277777777778)

    def test_yearfrac_basis_1(self):
        self.assertAlmostEqual(yearfrac(date(2008, 1, 1), date(2015, 4, 20), 1), 7.299110198)

    def test_yearfrac_basis_2(self):
        self.assertAlmostEqual(yearfrac(date(2008, 1, 1), date(2015, 4, 20), 2), 7.405555556)

    def test_yearfrac_basis_3(self):
        self.assertAlmostEqual(yearfrac(date(2008, 1, 1), date(2015, 4, 20), 3), 7.304109589)

    def test_yearfrac_basis_4(self):
        self.assertAlmostEqual(yearfrac(date(2008, 1, 1), date(2015, 4, 20), 4), 7.302777778)

    def test_yearfrac_inverted(self):
        self.assertAlmostEqual(yearfrac(date(2015, 4, 20), date(2008, 1, 1)), yearfrac(date(2008, 1, 1), date(2015, 4, 20)))    


class Test_Date(unittest.TestCase):
    def setup(self):
        pass

    def test_year_must_be_integer(self):
        with self.assertRaises(TypeError):
            date('2016', 1, 1)

    def test_month_must_be_integer(self):
        with self.assertRaises(TypeError):
            date(2016, '1', 1)

    def test_day_must_be_integer(self):
        with self.assertRaises(TypeError):
            date(2016, 1, '1')

    def test_year_must_be_positive(self):
        with self.assertRaises(ValueError):
            date(-1, 1, 1)

    def test_year_must_have_less_than_10000(self):
        with self.assertRaises(ValueError):
            date(10000, 1, 1)

    def test_result_must_be_positive(self):
        with self.assertRaises(ArithmeticError):
            date(1900, 1, -1)

    def test_not_stricly_positive_month_substracts(self):
        self.assertEqual(date(2009, -1, 1), date(2008, 11, 1))

    def test_not_stricly_positive_day_substracts(self):
        self.assertEqual(date(2009, 1, -1), date(2008, 12, 30))

    def test_month_superior_to_12_change_year(self):
        self.assertEqual(date(2009, 14, 1), date(2010, 2, 1))

    def test_day_superior_to_365_change_year(self):
        self.assertEqual(date(2009, 1, 400), date(2010, 2, 4))

    def test_year_for_29_feb(self):
        self.assertEqual(date(2008, 2, 29), 39507)

    def test_year_regular(self):
        self.assertEqual(date(2008, 11, 3), 39755)

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

    def test_negative_number_of_digits(self):
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


class Test_Countif(unittest.TestCase):
    def setUp(self):
        pass

    def test_argument_validity(self):
        with self.assertRaises(TypeError):
            countif(['e', 1], '>=1')

    def test_countif_strictly_superior(self):
        self.assertEqual(countif([7, 25, 13, 25], '>10'), 3)

    def test_countif_strictly_inferior(self):
        self.assertEqual(countif([7, 25, 13, 25], '<10'), 1)

    def test_countif_superior(self):
        self.assertEqual(countif([7, 10, 13, 25], '>=10'), 3)

    def test_countif_inferior(self):
        self.assertEqual(countif([7, 10, 13, 25], '<=10'), 2)

    def test_countif_different(self):
        self.assertEqual(countif([7, 10, 13, 25], '<>10'), 3)

    def test_countif_with_string_equality(self):
        self.assertEqual(countif([7, 'e', 13, 'e'], 'e'), 2)

    def test_countif_regular(self):
        self.assertEqual(countif([7, 25, 13, 25], 25), 2)


class Test_Countifs(unittest.TestCase): # more tests might be welcomed
    def setUp(self):
        pass

    def test_countifs_regular(self):
        self.assertEqual(countifs([7, 25, 13, 25], 25, [100, 102, 201, 20], ">100"), 1)

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
