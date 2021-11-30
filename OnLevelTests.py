import unittest
import numpy as np
import pandas as pd
import OnLevel as OnLevelClass
import OnLevelInput as OnLevelInputClass
import datetime
from TreatySlicer import CalendarTreatySlicer, PolicyTreatySlicer

class OnLevelTests(unittest.TestCase):
    
    input = OnLevelInputClass.OnLevelInput
    input.historicalPeriods = [[datetime.date(year, 1, 1), datetime.date(year, 12, 31)] 
        for year in range(2018, 2021)]
    input.prospectivePeriod = [datetime.date(2021, 1, 1), datetime.date(2021, 12, 31)]
        

    def test_one_rate_change_policy_to_policy(self):
        onlevel = OnLevelClass.OnLevel()
        self.input.rateChanges = pd.DataFrame (data= 
            {'date' : [datetime.date(2019,1,1)], 'rate' : [0.25] })
        self.input.historicalTreatySlicer = PolicyTreatySlicer
        self.input.prospectiveTreatySlicer = PolicyTreatySlicer
        actual = onlevel.Calculate(self.input)

        expected = [1.25, 1, 1]
        np.testing.assert_almost_equal(actual, expected, 7)

    
    def test_two_rate_changes_policy_to_policy(self):
        onlevel = OnLevelClass.OnLevel()
        
        self.input.rateChanges = pd.DataFrame (data= 
            {'date' : [datetime.date(2019,1,1), datetime.date(2019,7,1)],
             'rate' : [0.25, 0.1] })
        self.input.historicalTreatySlicer = PolicyTreatySlicer
        self.input.prospectiveTreatySlicer = PolicyTreatySlicer
        actual = onlevel.Calculate(self.input)

        expected = [1.375, 1.047187, 1]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_one_rate_change_calendar_to_calendar(self):
        onlevel = OnLevelClass.OnLevel()
        self.input.rateChanges = pd.DataFrame (data= 
            {'date' : [datetime.date(2019,1,1)], 'rate' : [0.25] })
        self.input.historicalTreatySlicer = CalendarTreatySlicer
        self.input.prospectiveTreatySlicer = CalendarTreatySlicer
        actual = onlevel.Calculate(self.input)

        expected = [1.25, 1.110773, 1]
        np.testing.assert_almost_equal(actual, expected, 7)

    
    def test_two_rate_changes_calendar_to_calendar(self):
        onlevel = OnLevelClass.OnLevel()
        
        self.input.rateChanges = pd.DataFrame (data= 
            {'date' : [datetime.date(2019,1,1), datetime.date(2019,7,1)],
             'rate' : [0.25, 0.1] })
        self.input.historicalTreatySlicer = CalendarTreatySlicer
        self.input.prospectiveTreatySlicer = CalendarTreatySlicer
        actual = onlevel.Calculate(self.input)

        expected = [1.375, 1.2047542, 1.0112685]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_one_rate_change_calendar_to_policy(self):
        onlevel = OnLevelClass.OnLevel()
        self.input.rateChanges = pd.DataFrame (data= 
            {'date' : [datetime.date(2019,1,1)], 'rate' : [0.25] })
        self.input.historicalTreatySlicer = CalendarTreatySlicer
        self.input.prospectiveTreatySlicer = PolicyTreatySlicer
        actual = onlevel.Calculate(self.input)

        expected = [1.25, 1.110773, 1]
        np.testing.assert_almost_equal(actual, expected, 7)

    
    def test_two_rate_changes_calendar_to_policy(self):
        onlevel = OnLevelClass.OnLevel()
        
        self.input.rateChanges = pd.DataFrame (data= 
            {'date' : [datetime.date(2019,1,1), datetime.date(2019,7,1)],
             'rate' : [0.25, 0.1] })
        self.input.historicalTreatySlicer = CalendarTreatySlicer
        self.input.prospectiveTreatySlicer = PolicyTreatySlicer
        actual = onlevel.Calculate(self.input)

        expected = [1.375, 1.2047542, 1.0112685]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_one_rate_change_policy_to_calendar(self):
        onlevel = OnLevelClass.OnLevel()
        self.input.rateChanges = pd.DataFrame (data= 
            {'date' : [datetime.date(2019,1,1)], 'rate' : [0.25] })
        self.input.historicalTreatySlicer = PolicyTreatySlicer
        self.input.prospectiveTreatySlicer = CalendarTreatySlicer
        actual = onlevel.Calculate(self.input)

        expected = [1.25, 1, 1]
        np.testing.assert_almost_equal(actual, expected, 7)

    
    def test_two_rate_changes_policy_to_calendar(self):
        onlevel = OnLevelClass.OnLevel()
        
        self.input.rateChanges = pd.DataFrame (data= 
            {'date' : [datetime.date(2019,1,1), datetime.date(2019,7,1)],
             'rate' : [0.25, 0.1] })
        self.input.historicalTreatySlicer = PolicyTreatySlicer
        self.input.prospectiveTreatySlicer = CalendarTreatySlicer
        actual = onlevel.Calculate(self.input)

        expected = [1.375, 1.047187, 1]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_two_rate_changes_with_duplicate_policy_to_policy(self):
        onlevel = OnLevelClass.OnLevel()
        
        self.input.rateChanges = pd.DataFrame (data= {
            'date' : [datetime.date(2019,1,1), datetime.date(2019,7,1), datetime.date(2019,7,1)], 
            'rate' : [0.25, 0.1, 0.15] })
        self.input.historicalTreatySlicer = PolicyTreatySlicer
        self.input.prospectiveTreatySlicer = PolicyTreatySlicer
        actual = onlevel.Calculate(self.input)
        
        expected = [1.58125, 1.1158666, 1]
        np.testing.assert_almost_equal(actual, expected, 7)

        
if __name__ == '__main__':
    unittest.main()