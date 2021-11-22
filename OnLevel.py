import pandas as pd
import datetime
import time
import random
from pprint import pprint


rateChanges = []
for x in range(0, 9):
    rateChanges.append( [datetime.datetime(2015, 1, 1).date() + datetime.timedelta(days=50 * (x+1)), random.uniform(0.00, 0.10)] )
#need group by and "sum"

print (" ")
print ("RATE CHANGES")
pprint (rateChanges)

factors = []
for rateChange in rateChanges:
    factors.append( [rateChange[0], 1 + rateChange[1]] )
print (" ")
print ("FACTORS")
pprint (factors)


cumultiveFactor = 1
reversedCumultiveFactors = []
for factor in reversed(factors):
    cumultiveFactor *= factor[1]
    reversedCumultiveFactors.append( [factor[0], cumultiveFactor] )


cumultiveFactors = []
for factor in reversed(reversedCumultiveFactors):
    cumultiveFactors.append(factor)

print (" ")
print ("CUMULATIVE FACTORS")
pprint (cumultiveFactors)

historicalPeriods = []
for year in range(2015, 2018):
    historicalPeriods.append( [datetime.datetime(year, 1, 1).date(), datetime.datetime(year, 12, 31).date()] )
    
prospectivePeriod = []
prospectivePeriod.append(["2022-01-01", "2022-12-31"])

#print (f"HistoricalPeriods: {historicalPeriods}")
#print (f"ProspectivePeriod: {prospectivePeriod}")
#print (f"RateChanges: {rateChanges}")
print (" ")


start = time.time()

earliestPolicyStartDate = historicalPeriods[0][0]
latestPolicyEndDate = prospectivePeriod[0][1]
df = pd.DataFrame(pd.date_range(start = earliestPolicyStartDate, end = latestPolicyEndDate), columns = ['PolicyStartDate'])
df["PolicyEndDate"] = df["PolicyStartDate"] + pd.DateOffset(years=1) + datetime.timedelta(days=-1)
df["DayCount"] = df["PolicyEndDate"] - df["PolicyStartDate"] + datetime.timedelta(days=1)



end = time.time()
#print(df)
print(end - start)