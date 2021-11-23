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
        rateChanges = pd.DataFrame(pd.date_range(start = "2015-01-01", periods = count), columns = ['Date'])
        rateChanges["Rate"] = .25
        
        historicalPeriods = []
        for year in range(2015, 2018):
            historicalPeriods.append( [datetime.datetime(year, 1, 1).date(), datetime.datetime(year, 12, 31).date()] )

        prospectivePeriod = []
        prospectivePeriod.append(["2022-01-01", "2022-12-31"])

        result = self.onlevel.Calculate(rateChanges, historicalPeriods, prospectivePeriod)
        self.assertEqual(result, 123)

    
    def test_Two_Rate_Changes(self):
        onlevel = OnLevelClass.OnLevel()
        
        # count = 2
        # rateChanges = pd.DataFrame(["2015-01-01", "2015-07-01"], columns = ['Date'])
        # for index in range(count):
        #     rateChanges['Date'].iloc[index] += datetime.timedelta(days=50 * (index+1))
        #     rateChanges["Rate"] = [0.25, 0.1]
    
        # historicalPeriods = []
        # for year in range(2015, 2018):
        #     historicalPeriods.append( [datetime.datetime(year, 1, 1).date(), datetime.datetime(year, 12, 31).date()] )

        # prospectivePeriod = []
        # prospectivePeriod.append(["2022-01-01", "2022-12-31"])
        
        result = self.onlevel.Calculate2()

        self.assertEqual(result, 123)

        
if __name__ == '__main__':
    unittest.main()