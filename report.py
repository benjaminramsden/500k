# -*- coding: utf-8 -*-
from datetime import datetime
from utils import validate_state
import logging

"""
Class for reports
"""
class Report(object):
    """docstring for Report."""
    def __init__(self, date, name, pid, report):
        super(Report, self).__init__()
        self.set_date(date, pid)
        self.name = name
        self.validate_id(pid)
        self.add_reports(report)

    def set_date(self, date, pid):
        if not date:
            logging.warning("No date for report ID {0}".format(pid))
            self._date = None
            self.historical = True
        elif len(date) < 8:
            self._date = date
            self.historical = True
        else:
            date_list = date.split(" ")
            date_list[1] = date_list[1].zfill(2)
            self._date = datetime.strptime(" ".join(date_list),
                                          '%a %d %b %Y at %H:%M')
            self.historical = False


    def get_month(self):
        if not self.historical:
            return self._date.month


    def get_year(self):
        if not self.historical:
            return self._date.year


    def add_reports(self, report):
        if len(report) < 2200:
            self.report = [report]
        else:
            logging.info("Splitting into several slides {0}".format(self.id))
            paragraphs = report.splitlines()
            if len(paragraphs) > 1:
                temp = paragraphs[0]
                self.report = [temp]
                counter = 1
                for para in paragraphs[1:]:
                    temp = temp + "\n" + para
                    if len(temp) > 2200:
                        break
                    self.report = [temp]
                    counter += 1
                self.report.extend(paragraphs[counter:])
            else:
                logging.warning("{0},{1} has no paragraphs".format(
                    self.id,
                    self._date))


    def validate_id(self, pid):
        if len(pid) != 6:
            raise NotImplementedError(
                'Missionary ID wrong length, ID: {0}'.format(pid))
        else:
            try:
                validate_state(pid[:2], abbreviation=True)
                self.id = pid
            except ValueError:
                print "ERROR: Invalid ID {0}".format(pid)
                raise NotImplementedError

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
