import configparser, os

## File setup

config = configparser.ConfigParser()

############# DEFAULT FILE SETUP ###############
config['OPTIONS'] = {}

config['Cell Borders'] = {'borders': 'true'}

config['Request Notes Resizing'] = {'resize_requestNotes': 'true'}
################################################

def writeConfigFile():
  if not os.path.exists('config.ini'):
    with open('config.ini', 'w') as configfile:
      config.write(configfile)

## File Reader
def readConfigFile():
  config = configparser.ConfigParser()
  config.read('config.ini')
  return config

def getBordersOption():
  return config.getboolean('Cell Borders', 'borders') 

def getRequestnotesOption():
  return config['Request Notes Resizing'].getboolean('resize_requestNotes')