"""
Class describing a missionary
"""
from datetime import datetime
from utils import validate_state
import logging


class Person(object):
    """Used to detail basic information on a person"""
    def __init__(self, first_name=None):
        super(Person, self).__init__()
        if first_name:
            self.first_name = first_name.title()


class Adult(Person):
    """Uses the information to build adult data."""
    def __init__(self, surname, first_name=None):
        super(Adult, self).__init__(first_name)
        self.surname = surname.rstrip('\r').title()


class Missionary(Adult):
    """
    Missionary class contains personal details and historical factfile and
    report information. Should allow initialisation just based on Missionary ID
    """
    def __init__(self, id, surname, first_name=None, pic=None):
        super(Missionary, self).__init__(surname, first_name)
        self.reports = {}
        self.children = {}
        self.pic = pic
        if not id or len(id) is not 6:
            raise NotImplementedError(
                "Invalid ID {0}, cannot process".format(id))
        else:
            try:
                validate_state(id[:2], True)
                self.id = id
            except ValueError:
                raise NotImplementedError(
                    "Invalid ID {0}, cannot process".format(id))


class Spouse(Adult):
    """
    Missionary's spouse information
    """
    def __init__(self, first_name, surname):
        super(Spouse, self).__init__(first_name, surname)


class Child(Person):
    """Uses the information available on the missionary's children."""
    def __init__(self, first_name, dob):
        super(Child, self).__init__(first_name)
        try:
            self._dob = datetime.strptime(dob, '%Y')
            self.age = datetime.now().year - self._dob.year
        except ValueError:
            logging.error("Invalid year of birth: {0}. Age will not be added".format(
                dob))
            self.age = None
