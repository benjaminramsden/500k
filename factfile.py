"""
This file contains the class Factfile representing all the information input
into the factfile uploader.
"""

class Factfile(object):
    """docstring for Factfile."""
    def __init__(self, first_name, pid, state, surname=None):
        super(Factfile, self).__init__()
        self.first_name = first_name
        self.state = state
        self.id = pid
        if surname:
            self.surname = surname

    def set_dob(day, month, year):
        """
        Store date of birth in datetime format
        """
        self._dob = datetime.strptime(
            " ".join(day.zfill(2),month.zfill(2),year),'%d %m %Y')

    def get_age():
        return datetime.now().year - self._dob.year

    def set_relationships(marital_status, spouse_fn=None, spouse_sn=None,
            children=None, no_deps=0):
        """
        Aims to encode all the relationships within the family.
        """
        self.marital_status = marital_status
        if spouse_fn:
            self.spouse_fn = spouse_fn
        if spouse_sn:
            self.spouse_sn = spouse_sn

    def set_languages(languages):
        self.languages = languages
