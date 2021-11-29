import unittest
import numpy as np
import pandas as pd
import OnLevel as OnLevelClass
import OnLevelInput as OnLevelInputClass
import datetime
from TreatySlicer import PolicyTreatySlicer

class OnLevelTests(unittest.TestCase):
    input = OnLevelInputClass.OnLevelInput
    input.historicalPeriods = [[datetime.date(year, 1, 1), datetime.date(year, 12, 31)] 
        for year in range(2018, 2021)]
    input.prospectivePeriod = [datetime.date(2021, 1, 1), datetime.date(2021, 12, 31)]
    input.historicalTreatySlicer = PolicyTreatySlicer
    input.prospectiveTreatySlicer = PolicyTreatySlicer
        

    def test_one_rate_change(self):
        onlevel = OnLevelClass.OnLevel()
        self.input.rateChanges = pd.DataFrame (data= 
            {'date' : [datetime.date(2019,1,1)], 'rate' : [0.25] })
        actual = onlevel.Calculate(self.input)

        expected = [1.25, 1, 1]
        np.testing.assert_almost_equal(actual, expected, 7)

    
    def test_two_rate_changes(self):
        onlevel = OnLevelClass.OnLevel()
        
        self.input.rateChanges = pd.DataFrame (data= 
            {'date' : [datetime.date(2019,1,1), datetime.date(2019,7,1)],
             'rate' : [0.25, 0.1] })
        actual = onlevel.Calculate(self.input)

        expected = [1.375, 1.047187, 1]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_two_rate_changes_with_duplicate(self):
        onlevel = OnLevelClass.OnLevel()
        
        self.input.rateChanges = pd.DataFrame (data= {
            'date' : [datetime.date(2019,1,1), datetime.date(2019,7,1), datetime.date(2019,7,1)], 
            'rate' : [0.25, 0.1, 0.15] })
        actual = onlevel.Calculate(self.input)
        
        expected = [1.58125, 1.1158666, 1]
        np.testing.assert_almost_equal(actual, expected, 7)

        
if __name__ == '__main__':
    unittest.main()