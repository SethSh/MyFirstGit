from pandas.core.frame import DataFrame
from TreatySlicer import BaseTreatySlicer

class OnLevelInput:
    rateChanges: DataFrame
    historicalPeriods: list
    prospectivePeriod: list
    historicalTreatySlicer: BaseTreatySlicer
    prospectiveTreatySlicer: BaseTreatySlicer