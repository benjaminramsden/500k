from datetime import datetime

"""
Class for reports
"""
class Report(object):
    """docstring for Report."""
    def __init__(self, date, name, pid, report):
        super(Report, self).__init__()
        if not date:
            print "ERROR: No date for report ID {0}".format(pid)
            self._date = None
        elif len(date) < 8:
            self._date = date
            self.historical = True
        else:
            date_list = date.split(" ")
            date_list[1] = date_list[1].zfill(2)
            self._date = datetime.strptime(" ".join(date_list),
                                          '%a %d %b %Y at %H:%M')
            self.historical = False
        self.name = name
        self.id = pid
        print "Report Lengths {0}".format(len(report))
        if len(report) < 3000:
            self.report = [report]
        else:
            paragraphs = report.split("\n")
            print "Number of paragraphs: {0}".format(len(paragraphs))
            if len(paragraphs) > 1:
                temp = paragraphs[0]
                self.report = [temp]
                print "self.report = {0}".format(self.report)
                counter = 1
                for para in paragraphs[1:]:
                    temp = temp + "\n" + para
                    print u"Length of combo {0}".format(len(temp))
                    if len(temp) > 3000:
                        break
                    self.report = [temp]
                    counter += 1
                self.report.extend(paragraphs[counter:])
                print "Report split: {0}".format(self.report)
            else:
                print "ERROR: SUPER LONG PARAGRAPH - MUST FIX"
                self.report = [report]

    def get_report_round(self):
        if not self._date:
            report_round = None
        elif self.historical:
            date_list = self._date.split(" ")
            year = date_list[0]
            round = date_list[1].lstrip("R")
            report_round = (round, year)
        else:
            if self._date.month in range(1,5):
                report_round = ("1", self._date.year)
            elif self._date.month in range(5,9):
                report_round = ("2", self._date.year)
            elif self._date.month in range(9,13):
                report_round = ("3", self._date.year)
            else:
                raise ValueError("Month value not in 1-12 bound")
        return report_round
