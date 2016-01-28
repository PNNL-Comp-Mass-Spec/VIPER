Syntax:
VIPER_MTS /A [/T:TraceLogLevel]
   or
VIPER_MTS ParameterFilePath.Par /R [/T:TraceLogLevel]
   or
VIPER_MTS /I:InputFilePath.xxx /N:IniFilePath.Ini /R [/T:TraceLogLevel]
   or
VIPER_MTS /G:FolderStartPath /O

Use of /A will initiate fully automated PRISM automation mode.  The 
database will be queried periodically to look for available jobs.

A parameter file can be used to list the input file path and JobNumber 
for auto analysis, along with other paths.  Example parameter file:

InputFilePath=C:\Inp\InFile.PEK
JobNumber=823
OutputFolderPath=C:\Out
IniFilePath=C:\Param\Settings.ini
LogFilePath=C:\Logs\LogFile.log

The file extension for the input file can be .Pek, .CSV, .mzXML, 
mzxml.xml, .mzData, or mzdata.xml Multiple PEK/CSV/mzXML/mzData files 
can be listed on the InputFilePath line, separating them using a 
vertical bar |.  In this case, also separate the JobNumber values 
using the vertical bar.  A speed advantage exists when processing 
multiple files in one call to this program, since the MT database data 
need only be loaded once.

Alternatively, use /I and /N for auto analysis without using a 
parameter file.  The /I switch specifies a .Pek, .CSV, etc. file to 
automatically process. If /N is missing, the options listed in 
VIPERSettings.ini will be used. /R is optional and means to not exit 
the program when done auto-processing (Remain Open). 

/T can be used to set the trace log level, for example /T:5  Higher 
numbers mean less logging /T:0 means off.  The default is off. 

/G instructs the program to examine the folder given by 
FolderStartPath and generate Index.html files for navigation, looking 
for Viper results folders containing Index.html files to determine the 
datafile names.  Unless /R is present, the program will exit when done.

