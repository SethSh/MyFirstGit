import unittest
import numpy as np
import pandas as pd
import OnLevel as OnLevelClass
import OnLevelInput as OnLevelInputClass
from datetime import date
from TreatySlicer import CalendarTreatySlicer, PolicyTreatySlicer
import random
import time

class OnLevelTests(unittest.TestCase):
    
    def setUp(self):
        self.startTime = time.time()
        
    def tearDown(self):
        t = time.time() - self.startTime
        print('%s: %.3f' % (self.id(), t))
        
    def test_three_rate_changes_policy_to_policy(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1), date(2019, 7, 1), date(2021, 7, 1)], 'rate' : [0.250000, 0.100000, 0.050000]})
        input.historicalTreatySlicer = PolicyTreatySlicer
        input.prospectiveTreatySlicer = PolicyTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.4096575, 1.0735818, 1.0252055]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_three_rate_changes_calendar_to_policy(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1), date(2019, 7, 1), date(2021, 7, 1)], 'rate' : [0.250000, 0.100000, 0.050000]})
        input.historicalTreatySlicer = CalendarTreatySlicer
        input.prospectiveTreatySlicer = PolicyTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.4096575, 1.2351206, 1.0367580]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_three_rate_changes_calendar_to_calendar(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1), date(2019, 7, 1), date(2021, 7, 1)], 'rate' : [0.250000, 0.100000, 0.050000]})
        input.historicalTreatySlicer = CalendarTreatySlicer
        input.prospectiveTreatySlicer = CalendarTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.3837792, 1.2124463, 1.0177253]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_three_rate_changes_policy_to_calendar(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1), date(2019, 7, 1), date(2021, 7, 1)], 'rate' : [0.250000, 0.100000, 0.050000]})
        input.historicalTreatySlicer = PolicyTreatySlicer
        input.prospectiveTreatySlicer = CalendarTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.3837792, 1.0538731, 1.0063849]
        np.testing.assert_almost_equal(actual, expected, 7)

    
    def test_two_rate_changes_policy_to_policy(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1), date(2021, 7, 1)], 'rate' : [0.250000, 0.050000]})
        input.historicalTreatySlicer = PolicyTreatySlicer
        input.prospectiveTreatySlicer = PolicyTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.2815068, 1.0252055, 1.0252055]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_two_rate_changes_policy_to_calendar(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1), date(2021, 7, 1)], 'rate' : [0.250000, 0.050000]})
        input.historicalTreatySlicer = PolicyTreatySlicer
        input.prospectiveTreatySlicer = CalendarTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.2579811, 1.0063849, 1.0063849]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_two_rate_changes_calendar_to_calendar(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1), date(2021, 7, 1)], 'rate' : [0.250000, 0.050000]})
        input.historicalTreatySlicer = CalendarTreatySlicer
        input.prospectiveTreatySlicer = CalendarTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.2579811, 1.1178651, 1.0063849]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_two_rate_changes_calendar_to_policy(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1), date(2021, 7, 1)], 'rate' : [0.250000, 0.050000]})
        input.historicalTreatySlicer = CalendarTreatySlicer
        input.prospectiveTreatySlicer = PolicyTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.2815068, 1.1387705, 1.0252055]
        np.testing.assert_almost_equal(actual, expected, 7)

    
    def test_one_rate_change_policy_to_policy(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1)], 'rate' : [0.250000]})
        input.historicalTreatySlicer = PolicyTreatySlicer
        input.prospectiveTreatySlicer = PolicyTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.2500000, 1.0000000, 1.0000000]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_one_rate_change_policy_to_calendar(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1)], 'rate' : [0.250000]})
        input.historicalTreatySlicer = PolicyTreatySlicer
        input.prospectiveTreatySlicer = CalendarTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.2500000, 1.0000000, 1.0000000]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_one_rate_change_calendar_to_calendar(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1)], 'rate' : [0.250000]})
        input.historicalTreatySlicer = CalendarTreatySlicer
        input.prospectiveTreatySlicer = CalendarTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.2500000, 1.1107730, 1.0000000]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_one_rate_change_calendar_to_policy(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(2018, 1, 1), date(2018, 12, 31)], [date(2019, 1, 1), date(2019, 12, 31)], [date(2020, 1, 1), date(2020, 12, 31)]]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(2019, 1, 1)], 'rate' : [0.250000]})
        input.historicalTreatySlicer = CalendarTreatySlicer
        input.prospectiveTreatySlicer = PolicyTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        expected = [1.2500000, 1.1107730, 1.0000000]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_three_rate_changes_calendar_to_calendar_many(self):
        input = OnLevelInputClass.OnLevelInput
        input.historicalPeriods = [[date(year, 1, 1), date(year, 12, 31)] for year in range(2000, 2020)]
        input.prospectivePeriod = [date(2021, 1, 1), date(2021, 12, 31)]
        input.rateChanges = pd.DataFrame (data = {'date' : [date(year, 1, 1) for year in range(1980,2020)], 'rate' : np.random.uniform(-.1,0.1,40).tolist()})
        input.historicalTreatySlicer = CalendarTreatySlicer
        input.prospectiveTreatySlicer = CalendarTreatySlicer

        onlevel = OnLevelClass.OnLevel()
        actual = onlevel.Calculate(input)

        # expected = [1.4096575, 1.4096575, 1.0735818, 1.0252055]
        # np.testing.assert_almost_equal(actual, expected, 7)




if __name__ == '__main__':
    suite = unittest.TestLoader().loadTestsFromTestCase(OnLevelTests)
    unittest.TextTestRunner(verbosity=0).run(suite)