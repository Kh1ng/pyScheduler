class Nerd:
    def __init__(self, name):
        self.name = name
        self.shiftCount = 15
        self.needOff = []

    def __repr__(self):
        return self.name