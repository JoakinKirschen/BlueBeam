Set args = WScript.Arguments

'// you can get url via parameter like line below
'//Url = args.Item(0)
Dim oShell : Set oShell = CreateObject("WScript.Shell")
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

Dim LabPathSource, LabPathDest, Temp, UserPath, PlotPath, PMPPath, PlotStyles, Plotters, sTemp, iTemp1, Url

Const TemporaryFolder = 2

dim myPath, a, filename

Wscript.Echo "Start downloading"

dim xHttp: Set xHttp = createobject("MSXML2.ServerXMLHTTP.3.0")

Url = "https://github.com/JoakinKirschen/BlueBeam/archive/master.zip"

a=split(Url,"/")
filename=a(ubound(a))

myPath = fso.GetSpecialFolder(TemporaryFolder) & "/" & filename 

xHttp.Open "GET", Url, False
xHttp.Send

'Wscript.Echo "Download-Status: " ^& xHttp.Status ^& " " ^& xHttp.statusText
 
If xHttp.Status = 200 Then
    Dim objStream
    set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1 'adTypeBinary
    objStream.Open
    objStream.Write xHttp.responseBody
    objStream.SaveToFile myPath,2
    objStream.Close
    set objStream = Nothing
End If
set xHttp=Nothing

Wscript.Echo "Start extracting"

'//Extract new version
Wscript.Echo "Extracting new version"
'The location of the zip file.
ExtractTo=fso.GetSpecialFolder(TemporaryFolder) '"C:\Test\"
ZipFile=ExtractTo & "\" & filename '"C:\Test.Zip"

Set objFolder = fso.GetFolder(ExtractTo)
Set colFiles = objFolder.Files
For Each objFile in colFiles
    If objFile.Attributes AND ReadOnly Then
        objFile.Attributes = objFile.Attributes XOR ReadOnly
    End If
Next

ExtractFolder = ExtractTo & "\BlueBeam-master"

If fso.FolderExists(ExtractFolder) Then
   fso.DeleteFolder ExtractFolder, True
End If
'If NOT fso.FolderExists(ExtractTo) Then
'   fso.CreateFolder(ExtractTo)
'End If

'Extract the contants of the zip file.
set objShell = CreateObject("Shell.Application")
set FilesInZip = objShell.NameSpace(ZipFile).items
objShell.NameSpace(ExtractTo).CopyHere(FilesInZip)
Set objShell = Nothing

CopySource=fso.GetSpecialFolder(TemporaryFolder) & "\BlueBeam-master\"
AppData = oShell.expandEnvironmentStrings("%APPDATA%")

Wscript.Echo "Copying new version"
'If fso.FolderExists(CopySource) Then 
'    fso.CopyFolder CopySource, AppData 
'End If

destfoldertemplate = AppData & "\Bluebeam Software\Revu\21\Templates\"

sdcopy "_SSD - A0 landscape.pdf" , destfoldertemplate
sdcopy "_SSD - A0-26 landscape.pdf" , destfoldertemplate
sdcopy "_SSD - A3 landscape.pdf" , destfoldertemplate
sdcopy "_SSD - A4 portrait.pdf" , destfoldertemplate
sdcopy "_SSD Bluebeam toolchest building template 20251126.pdf" , destfoldertemplate
sdcopy "_SSD stamp NOK.pdf" , destfoldertemplate
sdcopy "_SSD stamp OK, with comments.pdf" , destfoldertemplate
sdcopy "_SSD stamp OK.pdf" , destfoldertemplate

destfolderprofile = AppData & "\Bluebeam Software\Revu\21\"
sdcopy "Sweco STRU profile.bpx" , destfolderprofile


sub sdcopy (filename, destfolder)
   source = CopySource & filename
   dest = destfolder & filename
   If fso.FileExists(dest) Then
        'Check to see if the file is read-only
        If Not fso.GetFile(dest).Attributes And 1 Then 
            'The file exists and is not read-only.  Safe to replace the file.
            fso.CopyFile source, dest, True
        Else 
            'The file exists and is read-only.
            'Remove the read-only attribute
            fso.GetFile(dest).Attributes = fso.GetFile(dest).Attributes - 1
            'Replace the file
            fso.CopyFile source, dest, True
            'Reapply the read-only attribute
            fso.GetFile(dest).Attributes = fso.GetFile(dest).Attributes + 1
        End If
    Else
        'The file does not exist in the destination folder.  Safe to copy file to this folder.
        fso.CopyFile source, dest, True
    End If
End Sub
MsgBox("Update Successful")
