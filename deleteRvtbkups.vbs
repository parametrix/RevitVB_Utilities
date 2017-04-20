'	(C) Copyright 2016 by Francis Sebastian.
'	This program is free software; you can redistribute it and/or modify
'	it under the terms of the GNU General Public License as published by
'	the Free Software Foundation; either version 2 of the License, or
'	(at your option) any later version.
'
'	This program is distributed in the hope that it will be useful,
'	but WITHOUT ANY WARRANTY; without even the implied warranty of
'	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'	GNU General Public License for more details.
'
'	You should have received a copy of the GNU General Public License
'	along with this program; if not, write to the Free Software
'	Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
'	MA 02110-1301, USA.



Set oShell = CreateObject("WScript.Shell")
Set ofso = CreateObject("Scripting.FileSystemObject")
oShell.CurrentDirectory = ofso.GetParentFolderName(WScript.ScriptFullName)
'WScript.Echo oShell.CurrentDirectory
intConfirm = MsgBox("Do you want to delete all Revit rvt and rfa files in the folder: " & oShell.CurrentDirectory, vbYesNo, "Confirm Deletion")
if intConfirm = vbNo then WScript.Quit
dim dir: Set dir = ofso.GetFolder(oShell.CurrentDirectory)

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True   
objRegEx.IgnoreCase = True
objRegEx.Pattern = "(\.\d{4})+\.((rfa)|(rvt)$)"

Set re = New RegExp
re.Pattern = "^\d{4}$"
numFilesDeleted = 0
intAnswer = MsgBox("Delete files within sub folders?", vbYesNo, "Delete Files In SubFolder")
DeleteRevitBackupsRecursively(dir)
Wscript.Echo "Number of Files Deleted: " & numFilesDeleted

Sub DeleteRevitBackupsRecursively(Folder)
	'DeleteFiles(Folder)
	DeleteRevitFiles(Folder)
	if intAnswer = vbYes then
		For Each Subfolder in Folder.SubFolders
			Set objFolder = ofso.GetFolder(Subfolder.Path)
			DeleteRevitBackupsRecursively(objFolder)
		Next
	end if
End Sub

Sub DeleteRevitFiles(Folder)
	dim file: for each file in Folder.Files
		dim name: name = file.name
		dim extension: extension = ofso.GetExtensionName(name)
		if objRegEx.Test(name) then
			Dim objFile
			Set objFile = ofso.GetFile(file)
			objFile.Delete
			Set objFile = Nothing
			numFilesDeleted = numFilesDeleted+1
		end if
	next
End Sub