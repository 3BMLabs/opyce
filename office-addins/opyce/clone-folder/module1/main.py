#if you have not installed python:
#	go to the run section (click on the triangle bug icon on the left)
#	click 'Run and Debug'
#	 follow instructions.

from backend import main

#connect to Office app
opyce = main.Opyce()

#Cleanup the com reference. 
del opyce
