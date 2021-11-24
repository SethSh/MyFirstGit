from numpy.core.fromnumeric import size
import pandas as pd
import datetime
import time
import random
from pprint import pprint
import numpy as np
from pandas.core.frame import DataFrame

class OnLevel:
    def Calculate(self, rateChanges, historicalPeriods, prospectivePeriod):
        
        def AppendPolicyFactors(rateChanges, policies):
            if len(rateChanges) > 1:
                for index in range(len(rateChanges)-1):
                    lower = rateChanges['Date'].iloc[index]
                    upper = rateChanges['Date'].iloc[index + 1]    
                    f = rateChanges['CumulativeFactor'].iloc[index]
                    policies.loc[(policies["Start"] >= lower) & (policies["Start"] < upper), "CumulativeFactor"] = f

            if len(rateChanges) > 0:
                policies.loc[policies["Start"] <  rateChanges["Date"][0], "CumulativeFactor"] = rateChanges["CumulativeFactor"][0]
                policies.loc[policies["Start"] >= rateChanges["Date"].iloc[-1], "CumulativeFactor"] = rateChanges["CumulativeFactor"].iloc[-1]

        
        unsortedRateChanges = rateChanges
        unsortedRateChanges["Factor"] = unsortedRateChanges["Rate"] + 1
        
        rateChanges = unsortedRateChanges.groupby('Date').agg({'Factor': 'prod'}).sort_values(by="Date").reset_index()
        rateChanges["CumulativeFactor"] = rateChanges["Factor"].iloc[::-1].cumprod()
        
        earliestPolicyStartDate = historicalPeriods[0][0]
        latestPolicyEndDate = prospectivePeriod[0][1]
        policies = pd.DataFrame(pd.date_range(start = earliestPolicyStartDate, end = latestPolicyEndDate), columns = ['Start'])
        policies["PolicyEndDate"] = policies["Start"] + pd.DateOffset(years=1) + datetime.timedelta(days=-1)
        policies["DayCount"] = policies["PolicyEndDate"] - policies["Start"] + datetime.timedelta(days=1)
        policies["CumulativeFactor"] = 1

        AppendPolicyFactors(rateChanges, policies)
        return 123;    