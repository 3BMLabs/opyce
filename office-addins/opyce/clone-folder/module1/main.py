#if you have not installed python:
#	go to the run section (click on the triangle bug icon on the left)
#	click 'Run and Debug'
#	 follow instructions.

from backend import main

#Launch Excel and Open Workbook
opyce = main.Opyce()

#Run Macro
opyce.app.Application.Run("${opyce.workbook_name}.xlsm!Module1.testexcelmacro") 

#Cleanup the com reference. 
del opyce
