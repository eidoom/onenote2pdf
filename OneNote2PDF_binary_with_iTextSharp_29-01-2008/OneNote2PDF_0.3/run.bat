REM Sample batch file to export an entire notebook excluding some sections with Exclude command
set OneNote2PDF=.\OneNote2PDF.exe
set option=-TOCLevel 10 -ExportNotebook true  -Exclude "links,personal notes,projects"

%OneNote2PDF% %option% -Notebook "Shared Reference" -CacheFolder C:\Temp\OneNote -Output C:\Reading 
