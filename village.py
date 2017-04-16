"""
Class defining village information
"""
class Village(object):
    """docstring for Village."""
    def __init__(self, name, attendance, baptisms):
        super(Village, self).__init__()
        self.name = name
        self.attendance = attendance
        self.baptisms = baptisms
