#CAUTION! this is code is automatically generated!
import win32com.client

class Opyce:
	def __init__(self) -> None:
		#Launch Excel and Open Workbook
		self.app=win32com.client.GetActiveObject("$appname$.Application")

		$initialization$
		
	def __del__(self) -> None:
		del self.app