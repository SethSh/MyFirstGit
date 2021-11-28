import abc
import datetime

class BaseTreatySlicer(metaclass = abc.ABCMeta):
    @abc.abstractmethod
    def GetEarliestPolicyDate(start, lengthInYears):
        pass

    @abc.abstractmethod
    def GetPolicyFilter(policies, start, end, lengthInYears):
        pass

    @abc.abstractmethod
    def GetDays(policies, start, end):
        pass


class PolicyTreatySlicer(BaseTreatySlicer):
    def GetEarliestPolicyDate(start, lengthInYears):
        return start

    def GetPolicyFilter(policies, start, end, lengthInYears):
        return (policies.start_date >= start) & (policies.start_date <= end)

    def GetDays(policies, start, end):
        return (policies.end_date - policies.start_date).map(lambda x: x.days) + 1


class CalendarTreatySlicer(BaseTreatySlicer):
    def GetEarliestPolicyDate(start, lengthInYears):
        return start + datetime.timedelta(years=-1)

    def GetPolicyFilter(policies, start, end, lengthInYears):
        return (policies.start_date >= start + datetime.timedelta(years=-1)) & (policies.start_date <= end)

    def GetDays(policies, start, end):
        return max( min(policies.end, end) - max(policies.start, start) + 1, 0)
