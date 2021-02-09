import os, inspect

def locate():
	return os.path.abspath(inspect.stack()[-1][1])

def getParent():
	return os.path.dirname(locate())