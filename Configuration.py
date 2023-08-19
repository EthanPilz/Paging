import configparser, os

## File setup
config = configparser.ConfigParser()

############# DEFAULT FILE SETUP ###############
config['------------- Program Options -------------'] = {'All of these options ONLY accept' : '\'true\' or \'false\' \n If you think you messed up the configuration file, just delete it and re-run the program \n it will generate a new one with default values. \n \n You can view my source code at \n https://github.com/EthanPilz/UW-Odegaard-Paging-Automator'}

config['Cell Borders'] = {'borders': 'true'}

config['Request Notes Resizing'] = {'resize_requestNotes': 'true'}

config['Automatically print'] = {'print': 'true'}

config['Start at downloads folder for spreadsheet file select'] = {'downloads_start': 'true'}

config['Output new file to downloads folder'] = {'output_to_downloads': 'true'}
################################################

## Write the file if it doesn't exist
def writeConfigFile():
  if not os.path.exists('config.yaml'):
    with open('config.yaml', 'w') as configfile:
      config.write(configfile)

## File Reader
def readConfigFile():
  config = configparser.ConfigParser()
  config.read('config.yaml')
  return config

def getBordersOption(config):
  return config.getboolean('Cell Borders', 'borders') 

def getRequestnotesOption(config):
  return config.getboolean('Request Notes Resizing', 'resize_requestNotes') 

def getPrintOption(config):
  return config.getboolean('Automatically print', 'print')

def getInputStartOption(config):
  return config.getboolean('Start at downloads folder for spreadsheet file select', 'downloads_start')

def getOutputFolderOption(config):
  return config.getboolean('Output new file to downloads folder', 'output_to_downloads')