class Village(object):
    """
    Class defining village information
    """
    def __init__(self, name, attendance, baptisms):
        super(Village, self).__init__()
        self.name = name
        self.attendance = attendance
        self.baptisms = baptisms
