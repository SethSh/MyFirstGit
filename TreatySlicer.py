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
        return start.replace(year = start.year - lengthInYears) + datetime.timedelta(days=1)

    def get_policy_filter(policies, start, end, lengthInYears):
        newStart = start.replace(year = start.year - lengthInYears) + datetime.timedelta(days=1)
        return (policies.start_date >= newStart) & (policies.start_date <= end)

    def get_days(policies, start, end):
        days = policies.end_date.apply(lambda x: min(x, end)) - policies.start_date.apply(lambda x : max(x, start))
        return days.map(lambda x: x.days) + 1
