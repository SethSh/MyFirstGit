import numpy as np
import pandas as pd
import datetime


class OnLevel:

    def Calculate(self, input):
        preAggregatedRateChanges = input.rateChanges.copy(deep=True)
        preAggregatedRateChanges['incr_factor'] = preAggregatedRateChanges.rate + 1
        rateChanges = preAggregatedRateChanges.groupby('date').agg({'incr_factor': 'prod'}).reset_index()
        

        earliestPolicyStart = input.historicalTreatySlicer.GetEarliestPolicyDate(min(input.historicalPeriods[0]), 1)
        latestPolicyEndDate = input.prospectivePeriod[1]
        
        policies = pd.DataFrame(pd.date_range(start = earliestPolicyStart, end = latestPolicyEndDate), columns = ['start_date'])
        policies['end_date'] = policies.start_date + pd.DateOffset(years=1) + datetime.timedelta(days=-1)

        policies.start_date = policies.start_date.apply(lambda x: x.date())
        policies.end_date = policies.end_date.apply(lambda x: x.date())
        
        self.__CreatePolicyCumulativeFactors(rateChanges, policies)

        historicalFactors = []
        for period in input.historicalPeriods:
            totalDays = totalWeightedDays = 0
            treatyStart, treatyEnd = period 
            policyFilter = input.historicalTreatySlicer.GetPolicyFilter(policies, treatyStart, treatyEnd, 1)                       
            subset = policies.loc[policyFilter]

            factor = 1
            if any(subset):
                days = input.historicalTreatySlicer.GetDays(subset, treatyStart, treatyEnd)                       
                totalDays += days.sum()
                totalWeightedDays += (days * subset.cumul_factor).sum()
                factor = totalWeightedDays/totalDays if totalDays else 1
            historicalFactors.append(factor)

        treatyStart, treatyEnd = input.prospectivePeriod
        policyFilter = input.historicalTreatySlicer.GetPolicyFilter(policies, treatyStart, treatyEnd, 1)                       
        subset = policies.loc[policyFilter]
        days = input.prospectiveTreatySlicer.GetDays(subset, treatyStart, treatyEnd)                       
        totalDays = days.sum()
        totalWeightedDays = (days * subset.cumul_factor).sum()
        prospectiveFactor = totalWeightedDays/totalDays if totalDays else 1
 
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