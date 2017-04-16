from datetime import datetime

"""
Class for reports
"""
class Report(object):
    """docstring for Report."""
    def __init__(self, date, name, pid, report):
        super(Report, self).__init__()
        if len(date) < 8:
            self._date = date
            self.historical = True
        else:
            date_list = date.split(" ")
            date_list[1] = date_list[1].zfill(2)
            print date_list[1]
            self._date = datetime.strptime(" ".join(date_list),
                                          '%a %d %b %Y at %H:%M')
            self.historical = False
        self.name = name
        self.id = pid
        self.report = report

    def get_report_round(self):
        if self.historical:
            date_list = self._date.split(" ")[-1]
            year = date_list[0]
            round = date_list[1].lstrip("R")
            report_round = round + "/" + year
        else:
            if self._date.month in range(1,5):
                report_round = "1/"+str(self._date.year)
            elif self._date.month in range(5,9):
                report_round = "2/"+str(self._date.year)
            elif self._date.month in range(9,13):
                report_round = "3/"+str(self._date.year)
            else:
                raise ValueError("Month value not in 1-12 bound")
        return report_round
