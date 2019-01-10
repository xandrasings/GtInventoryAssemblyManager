import datetime

def getDateTime():
	return datetime.datetime.now()

def getDateTimeAsFilePathSegment():
	return getDateTime().strftime("%Y-%m-%d_%H%M")