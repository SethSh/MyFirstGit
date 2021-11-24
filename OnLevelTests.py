import unittest
import pandas as pd
import OnLevel as OnLevelClass
import datetime

class OnLevelTests(unittest.TestCase):
    historicalPeriods = [[datetime.date(year, 1, 1), datetime.date(year, 12, 31)] for year in range(2015, 2018)]
    prospectivePeriod = [datetime.date(2022, 1, 1), datetime.date(2022, 12, 31)]

    def test_One_Rate_Change(self):
        onlevel = OnLevelClass.OnLevel()
        
        rateChanges = pd.DataFrame (data= {'Date' : [datetime.date(2015,1,1)], 'Rate' : [0.25] })
        result = onlevel.Calculate(rateChanges, self.historicalPeriods, self.prospectivePeriod)
        self.assertEqual(result, 123)

    
    def Test_Two_Rate_Changes(self):
        onlevel = OnLevelClass.OnLevel()
        
        rateChanges = pd.DataFrame (data= {'Date' : [datetime.date(2015,1,1), datetime.datetime.date(2015,7,1)],'Rate' : [0.25, 0.1] })
        result = onlevel.Calculate(rateChanges, self.historicalPeriods, self.prospectivePeriod)
        self.assertEqual(result, 123)


    def Test_Two_Rate_Changes_With_Duplicate(self):
        onlevel = OnLevelClass.OnLevel()
        
        rateChanges = pd.DataFrame (data= {'Date' : [datetime.date(2015,1,1), datetime.date(2015,7,1), datetime.date(2015,7,1)], 'Rate' : [0.25, 0.1, 0.15] })            
        result = onlevel.Calculate(rateChanges, self.historicalPeriods, self.prospectivePeriod)
        self.assertEqual(result, 123)

        
if __name__ == '__main__':
    unittest.main()