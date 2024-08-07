opyce_shared: a folder containing shared global variables
opyce: contains the ribbon.

#adding an addin
to add an office app to these add-ins, create a new VSTO add-in project in the same solution.
add a reference to opyce.
then copy the code from another add-in and modify to your needs.

adding the ribbon:
https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.tools.word.documentbase.createribbonextensibilityobject?view=vsto-2022
