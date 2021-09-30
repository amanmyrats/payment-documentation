percentage=0

def resetpercentage():
    global percentage
    percentage=0

def updatepercentage(**kwargs):
    """ ['newpercentage'] """
    global percentage
    percentage+=kwargs['newpercentage']

def getpercentage():
    global percentage
    return percentage