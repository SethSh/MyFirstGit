from numpy.core.fromnumeric import size
import pandas as pd
import datetime
import time
import random
from pprint import pprint
import numpy as np

class OnLevel:
    def Calculate(self, rateChanges, historicalPeriods, prospectivePeriod):
        def AppendPolicyFactors(rateChanges, Policies):
            for index in range(len(rateChanges)-1):
                lower = rateChanges['Date'].iloc[index]
                upper = rateChanges['Date'].iloc[index + 1]    
                f = rateChanges['CumulativeFactor'].iloc[index]
                Policies.loc[(Policies["Start"] >= lower) & (Policies["Start"] < upper), "CumulativeFactor"] = f

            Policies.loc[Policies["Start"] <  rateChanges["Date"][0], "CumulativeFactor"] = rateChanges["CumulativeFactor"][0]
            Policies.loc[Policies["Start"] >= rateChanges["Date"].iloc[-1], "CumulativeFactor"] = rateChanges["CumulativeFactor"].iloc[-1]

        unsortedRateChanges = rateChanges
        unsortedRateChanges["Factor"] = unsortedRateChanges["Rate"] + 1

        rateChanges = unsortedRateChanges.sort_values(by="Date")
        rateChanges["CumulativeFactor"] = rateChanges["Factor"].iloc[::-1].cumprod()
        #need group by and "sum"

        earliestPolicyStartDate = historicalPeriods[0][0]
        latestPolicyEndDate = prospectivePeriod[0][1]
        df = pd.DataFrame(pd.date_range(start = earliestPolicyStartDate, end = latestPolicyEndDate), columns = ['Start'])
        df["PolicyEndDate"] = df["Start"] + pd.DateOffset(years=1) + datetime.timedelta(days=-1)
        df["DayCount"] = df["PolicyEndDate"] - df["Start"] + datetime.timedelta(days=1)
        df["CumulativeFactor"] = 1

        AppendPolicyFactors(rateChanges, df)
        return 123;    

    def Calculate2(self):
        return 123;    