import unittest
from OnLevel import OnLevel
import datetime
import pandas as pd

class OnLevelTests(unittest.TestCase):
    
    def Test_One_Rate_Change(self):
        onlevel = OnLevel()
        
        count = 1
        rateChanges = pd.DataFrame(pd.date_range(start = "2015-01-01", periods = count), columns = ['Date'])
        rateChanges["Rate"] = .25
        
        self.assertEqual(widget.size(), (50, 50))

    
    def Test_Two_Rate_Changes(self):
        onlevel = OnLevel()
        
        count = 2
        rateChanges = pd.DataFrame(["2015-01-01", "2015-07-01"], columns = ['Date'])
        for index in range(count):
            rateChanges['Date'].iloc[index] += datetime.timedelta(days=50 * (index+1))
            rateChanges["Rate"] = [0.25, 0.1]
    
        self.assertEqual(widget.size(), (50, 50))
    