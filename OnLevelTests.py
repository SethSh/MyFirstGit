import unittest
import numpy as np
import pandas as pd
import OnLevel as OnLevelClass
import datetime

class OnLevelTests(unittest.TestCase):
    historicalPeriods = [[datetime.date(year, 1, 1), datetime.date(year, 12, 31)] for year in range(2015, 2018)]
    prospectivePeriod = [datetime.date(2022, 1, 1), datetime.date(2022, 12, 31)]
        

    def test_one_rate_change(self):
        onlevel = OnLevelClass.OnLevel()
        
        rateChanges = pd.DataFrame (data= {'date' : [datetime.date(2016,1,1)], 'rate' : [0.25] })
        actual = onlevel.Calculate(rateChanges, self.historicalPeriods, self.prospectivePeriod)

        expected = [1.25, 1, 1]
        np.testing.assert_almost_equal(actual, expected, 7)

    
    def test_two_rate_changes(self):
        onlevel = OnLevelClass.OnLevel()
        
        rateChanges = pd.DataFrame (data= {'date' : [datetime.date(2016,1,1), datetime.date(2016,7,1)],'rate' : [0.25, 0.1] })
        actual = onlevel.Calculate(rateChanges, self.historicalPeriods, self.prospectivePeriod)
        
        expected = [1.375, 1.04736865, 1]
        np.testing.assert_almost_equal(actual, expected, 7)


    def test_two_rate_changes_with_duplicate(self):
        onlevel = OnLevelClass.OnLevel()
        
        rateChanges = pd.DataFrame (data= {'date' : [datetime.date(2015,1,1), datetime.date(2015,7,1), datetime.date(2015,7,1)], 'rate' : [0.25, 0.1, 0.15] })            
        actual = onlevel.Calculate(rateChanges, self.historicalPeriods, self.prospectivePeriod)
        
        expected = [1.11586659, 1, 1]
        np.testing.assert_almost_equal(actual, expected, 7)

        
if __name__ == '__main__':
    unittest.main()