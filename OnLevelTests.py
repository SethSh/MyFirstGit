import unittest
import datetime
import pandas as pd
import OnLevel as OnLevelClass

class OnLevelTests(unittest.TestCase):
    
    # def test_upper(self):
    #     self.assertEqual('foo'.upper(), 'FOO')
        
    def test_One_Rate_Change(self):
        onlevel = OnLevelClass.OnLevel()
        
        count = 1
        rateChanges = pd.DataFrame (data=
                        {
                            'Date' : ["2015-01-01"],
                            'Rate' : [0.25]
                        })
        
        historicalPeriods = []
        for year in range(2015, 2018):
            historicalPeriods.append( [datetime.datetime(year, 1, 1).date(), datetime.datetime(year, 12, 31).date()] )

        prospectivePeriod = []
        prospectivePeriod.append(["2022-01-01", "2022-12-31"])

        result = onlevel.Calculate(rateChanges, historicalPeriods, prospectivePeriod)
        self.assertEqual(result, 123)

    
    def test_Two_Rate_Changes(self):
        onlevel = OnLevelClass.OnLevel()
        
        count = 2
        rateChanges = pd.DataFrame (data=
                        {
                            'Date' : ["2015-01-01", "2015-07-01"],
                            'Rate' : [0.25, 0.1]
                        })
            
        historicalPeriods = []
        for year in range(2015, 2018):
            historicalPeriods.append( [datetime.datetime(year, 1, 1).date(), datetime.datetime(year, 12, 31).date()] )

        prospectivePeriod = []
        prospectivePeriod.append(["2022-01-01", "2022-12-31"])
        
        result = onlevel.Calculate(rateChanges, historicalPeriods, prospectivePeriod)
        self.assertEqual(result, 123)


    def test_Two_Rate_Changes_With_Duplicate(self):
        onlevel = OnLevelClass.OnLevel()
        
        count = 2
        rateChanges = pd.DataFrame (data=
                        {
                            'Date' : ["2015-01-01", "2015-07-01", "2015-07-01"],
                            'Rate' : [0.25, 0.1, 0.15]
                        })
            
        historicalPeriods = []
        for year in range(2015, 2018):
            historicalPeriods.append( [datetime.datetime(year, 1, 1).date(), datetime.datetime(year, 12, 31).date()] )

        prospectivePeriod = []
        prospectivePeriod.append(["2022-01-01", "2022-12-31"])
        
        result = onlevel.Calculate(rateChanges, historicalPeriods, prospectivePeriod)
        self.assertEqual(result, 123)

        
if __name__ == '__main__':
    unittest.main()