from pkg_resources import WorkingSet,DistributionNotFound
from setuptools.command.easy_install import main as installer
import platform
import subprocess

workingSet=WorkingSet()
dependencies={
    'openpyxl':'2.5.0a1',
    'apscheduler':'3.3.1',
    'pyttsx3':'2.6',
    #'twisted':'16.4.0' # just for testing purpose not required delte it later
}

    
try:
   for key,value in dependencies.items():
            dep=workingSet.require('{}>={}'.format(key,value))
except:
   try:
      if platform.system()=='Windows':
         installer(['{}>={}'.format(key,value)])
      elif platform.system()=='Linux':
         installer(['{}>={}'.format(key,value)]) # convert this to pip later
      print('{} {} has been installed successfully'.format(key,value))
   except:
      print('Dependecy installation error when installing {}>={}'.format(key,value))
      pass
   pass
