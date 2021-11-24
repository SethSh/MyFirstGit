import pandas as pd
import datetime

class OnLevel:

    def Calculate(self, rateChanges, historicalPeriods, prospectivePeriod):
        unsortedRateChanges = rateChanges
        unsortedRateChanges['Factor'] = unsortedRateChanges['Rate'] + 1
        rateChanges = unsortedRateChanges.groupby('Date').agg({'Factor': 'prod'}).sort_values(by='Date').reset_index()
        rateChanges['CumulativeFactor'] = rateChanges['Factor'].iloc[::-1].cumprod()
        
        earliestPolicyStartDate = historicalPeriods[0][0]
        latestPolicyEndDate = prospectivePeriod[1]
        
        policies = pd.DataFrame(pd.date_range(start = earliestPolicyStartDate, end = latestPolicyEndDate), columns = ['Start'])
        policies['End'] = policies['Start'] + pd.DateOffset(years=1) + datetime.timedelta(days=-1)

        policies.Start = policies.Start.apply(lambda x: x.date())
        policies.End = policies.End.apply(lambda x: x.date())
        
        policies['DayCount'] = policies['End'] - policies['Start'] + datetime.timedelta(days=1)
        self.__AppendPolicyCumulativeFactors(rateChanges, policies)

        for period in historicalPeriods:
            days = 0
            weightedDays = 0
            treatyStart, treatyEnd = period
            print(treatyStart)
            
            subset = policies.loc[(policies.Start >= treatyStart) & (policies.End < treatyEnd)]
            days += subset['DayCount'].sum()
            weightedDays += (subset['DayCount'] * subset['CumulativeFactor']).sum()

        return 123;    


    def __AppendPolicyCumulativeFactors(self, rateChanges, policies):
        policies['CumulativeFactor'] = 1
        if len(rateChanges) > 1:
            for index in range(len(rateChanges)-1):
                lower = rateChanges.Date.iloc[index]
                upper = rateChanges.Date.iloc[index + 1]    
                f = rateChanges['CumulativeFactor'].iloc[index]
                policies.loc[(policies.Start >= lower) & (policies.Start < upper), 'CumulativeFactor'] = f

        if len(rateChanges) > 0:
            policies.loc[policies.Start <  rateChanges.Date[0], 'CumulativeFactor'] = rateChanges.CumulativeFactor[0]
            policies.loc[policies.Start >= rateChanges.Date.iloc[-1], 'CumulativeFactor'] = rateChanges.CumulativeFactor.iloc[-1]