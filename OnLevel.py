import numpy as np
import pandas as pd
import datetime


class OnLevel:

    def Calculate(self, input):
        preAggregatedRateChanges = input.rateChanges
        preAggregatedRateChanges['incr_factor'] = preAggregatedRateChanges.rate + 1
        rateChanges = preAggregatedRateChanges.groupby('date').agg({'incr_factor': 'prod'}).reset_index()
        
        earliestPolicyStartDate = input.historicalPeriods[0][0]
        latestPolicyEndDate = input.prospectivePeriod[1]
        
        policies = pd.DataFrame(pd.date_range(start = earliestPolicyStartDate, end = latestPolicyEndDate), columns = ['start_date'])
        policies['end_date'] = policies.start_date + pd.DateOffset(years=1) + datetime.timedelta(days=-1)

        policies.start_date = policies.start_date.apply(lambda x: x.date())
        policies.end_date = policies.end_date.apply(lambda x: x.date())
        
        policies['days'] = (policies.end_date - policies.start_date).map(lambda x: x.days) + 1
        self.__CreatePolicyCumulativeFactors(rateChanges, policies)

        historicalFactors = []
        for period in input.historicalPeriods:
            days = weightedDays = 0
            treatyStart, treatyEnd = period 
            policyFilter = self.__GetFilter(policies, treatyStart, treatyEnd)                       
            subset = policies.loc[policyFilter]

            factor = 1
            if any(subset):
                days += subset.days.sum()
                weightedDays += (subset.days * subset.cumul_factor).sum()
                factor = weightedDays/days if days else 1
            historicalFactors.append(factor)

        treatyStart, treatyEnd = input.prospectivePeriod
        policyFilter = self.__GetFilter(policies, treatyStart, treatyEnd)                       
        subset = policies.loc[policyFilter]
        days = subset.days.sum()
        weightedDays = (subset.days * subset.cumul_factor).sum()
        prospectiveFactor = weightedDays/days if days else 1
 
        onlevelFactors = np.multiply(np.reciprocal(historicalFactors), prospectiveFactor)
        return onlevelFactors.tolist()


    def __GetFilter(self, policies, treatyStart, treatyEnd):
        policyFilter = (policies.start_date >= treatyStart) & (policies.start_date <= treatyEnd)
        return policyFilter    


    def __CreatePolicyCumulativeFactors(self, rateChanges, policies):
        policies['cumul_factor'] = 1

        for date, factor in zip(rateChanges.date, rateChanges.incr_factor):
            policyFilter = policies.start_date >= date
            policies.loc[policyFilter,'cumul_factor'] *= factor