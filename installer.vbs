Set args = WScript.Arguments
 
'// you can get url via parameter like line below
'//Url = args.Item(0)
Dim oShell : Set oShell = CreateObject("WScript.Shell")
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

Dim LabPathSource, LabPathDest, Temp, UserPath, PlotPath, PMPPath, PlotStyles, Plotters, sTemp, iTemp1, Url

Const TemporaryFolder = 2

dim myPath, a, filename
dim xHttp: Set xHttp = createobject("MSXML2.ServerXMLHTTP.3.0")

Url = "https://github.com/JoakinKirschen/BleuBeam/archive/master.zip"

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

Wscript.Echo "Download complete"

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

ExtractFolder = ExtractTo & "\lab2012-master"

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

CopySource=fso.GetSpecialFolder(TemporaryFolder) & "\lab2012-master"

Wscript.Echo "Copying new version"
If fso.FolderExists(CopySource) Then 
    fso.CopyFolder CopySource, LabPathDest 
End If

MsgBox("Update Successful")
