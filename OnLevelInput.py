from dataclasses import dataclass

from pandas.core.frame import DataFrame

class OnLevelInput:
    rateChanges: DataFrame
    historicalPeriods: list
    prospectivePeriod: list