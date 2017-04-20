' Path to Central File
Const SourceFile = "D:\_DEMO_DATASETS\_2017_Datasets\MedicalCenter\56750_Arch_2017\56750_M_Systems.rvt"
' Path to local folder
Const DestinationFolder = "C:\RevitLocal\"
'Path to Revit Executable
Const RevitEXE = """C:\Program Files\Autodesk\Revit 2017\Revit.exe"""

Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

    Dim sourceFolder
    sourceFolder = fso.GetParentFolderName(SourceFile)

    Dim localFileName
    localFileName = fso.GetBaseName(SourceFile) & "_LOCAL.rvt"

    Dim destinationFile
    destinationFile = DestinationFolder & localFileName

    ' from: http://stackoverflow.com/questions/1260740/copy-a-file-from-one-folder-to-another-using-vbscripting
    'Check to see if the file already exists in the destination folder
    If fso.FileExists(destinationFile) Then
        'Check to see if the file is read-only
        If Not fso.GetFile(destinationFile).Attributes And 1 Then 
            'The file exists and is not read-only.  Safe to replace the file.
            fso.CopyFile SourceFile, destinationFile, True
        Else 
            'The file exists and is read-only.
            'Remove the read-only attribute
            fso.GetFile(destinationFile).Attributes = fso.GetFile(destinationFile).Attributes - 1
            'Replace the file
            fso.CopyFile SourceFile, destinationFile, True
            'Reapply the read-only attribute
            fso.GetFile(destinationFile).Attributes = fso.GetFile(destinationFile).Attributes + 1
        End If
    Else
        'The file does not exist in the destination folder.  Safe to copy file to this folder.
        fso.CopyFile SourceFile, destinationFile, True
    End If

    ' Launching Revit and Opening File
    Dim WshShell
    Set WshShell = WScript.CreateObject("WScript.Shell")
    WshShell.Run RevitEXE & " " & destinationFile

Set fso = Nothing