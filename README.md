# Plant-3D-Batch-exporttoautocad

This tool (proof-of-concept, sample code, use at own risk) is exporting all project files as AutoCAD. It keeps the folder structure for the export. There are two commands, one for P&ID project files and another one for Plant 3D project files. "related files" are not included in the export, as they should be already plain AutoCAD.

How to use the script:

BatchExportToAutoCAD – exports 3d files, destination selection by file dialog
PnIdBatchExportToAutoCAD – exports PnID files, destination selection by file dialog

BatchExportToAutoCADHL – exports 3d files, destination by text input, this is good for batching the command
PnIdBatchExportToAutoCADHL – exports PnID files, destination by text input, this is good for batching the command

Critial code parts for this tool are based on:

https://www.keanw.com/2014/03/autocad-2015-calling-commands.html
https://adndevblog.typepad.com/autocad/2012/05/when-to-lock-the-document.html


