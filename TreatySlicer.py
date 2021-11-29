import abc
import datetime

class BaseTreatySlicer(metaclass = abc.ABCMeta):
    @abc.abstractmethod
    def get_first_policy_date(start, lengthInYears):
        pass

    @abc.abstractmethod
    def get_policy_filter(policies, start, end, lengthInYears):
        pass

    @abc.abstractmethod
    def get_days(policies, start, end):
        pass


class PolicyTreatySlicer(BaseTreatySlicer):
    def get_first_policy_date(start, lengthInYears):
        return start

    def get_policy_filter(policies, start, end, lengthInYears):
        return (policies.start_date >= start) & (policies.start_date <= end)

    def get_days(policies, start, end):
        return (policies.end_date - policies.start_date).map(lambda x: x.days) + 1


class CalendarTreatySlicer(BaseTreatySlicer):
    def get_first_policy_date(start, lengthInYears):
        return start + datetime.timedelta(years=-1)

    def get_policy_filter(policies, start, end, lengthInYears):
        newStart = start + datetime.timedelta(years=-1, days=1)
        return (policies.start_date >= newStart) & (policies.start_date <= end)

    def get_days(policies, start, end):
        return min(policies.end, end) - max(policies.start, start) + 1
