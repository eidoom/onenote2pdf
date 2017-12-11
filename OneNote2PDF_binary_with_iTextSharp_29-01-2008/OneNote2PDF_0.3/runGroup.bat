REM Sample batch file to export only one section/section group (and its children)

set OneNote2PDF=.\OneNote2PDF.exe
set option=-ExportSection "learning/programming/asp.net" -ShowTOC true
%OneNote2PDF% %option% -Notebook "Shared Reference" -CacheFolder C:\Temp\OneNote -Output C:\Reading