import numpy as np
import pandas as pd
import datetime

class OnLevel:

    def Calculate(self, inputRateChanges, historicalPeriods, prospectivePeriod):
        preAggregatedRateChanges = inputRateChanges
        preAggregatedRateChanges['incr_factor'] = preAggregatedRateChanges.rate + 1
        rateChanges = preAggregatedRateChanges.groupby('date').agg({'incr_factor': 'prod'}).reset_index()
        
        earliestPolicyStartDate = historicalPeriods[0][0]
        latestPolicyEndDate = prospectivePeriod[1]
        
        policies = pd.DataFrame(pd.date_range(start = earliestPolicyStartDate, end = latestPolicyEndDate), columns = ['start_date'])
        policies['end_date'] = policies.start_date + pd.DateOffset(years=1) + datetime.timedelta(days=-1)

        policies.start_date = policies.start_date.apply(lambda x: x.date())
        policies.end_date = policies.end_date.apply(lambda x: x.date())
        
        policies['days'] = (policies.end_date - policies.start_date).map(lambda x: x.days) + 1
        self.__CreatePolicyCumulativeFactors(rateChanges, policies)

        historicalFactors = []
        for period in historicalPeriods:
            days = weightedDays = 0
            treatyStart, treatyEnd = period                        
            subset = policies.loc[(policies.start_date >= treatyStart) & (policies.start_date <= treatyEnd)]

            factor = 1
            if any(subset):
                days += subset.days.sum()
                weightedDays += (subset.days * subset.cumul_factor).sum()
                factor = weightedDays/days if days else 1
            historicalFactors.append(factor)

        treatyStart, treatyEnd = prospectivePeriod
        subset = policies.loc[(policies.start_date >= treatyStart) & (policies.start_date <= treatyEnd)]
        days = subset.days.sum()
        weightedDays = (subset.days * subset.cumul_factor).sum()
        prospectiveFactor = weightedDays/days if days else 1
 
        onlevelFactors = np.multiply(np.reciprocal(historicalFactors), prospectiveFactor)
        return onlevelFactors.tolist();    


    def __CreatePolicyCumulativeFactors(self, rateChanges, policies):
        policies['cumul_factor'] = 1

        for d, f in zip(rateChanges.date, rateChanges.incr_factor):
            tf = policies.start_date >= d
            policies.loc[tf,'cumul_factor'] *= f