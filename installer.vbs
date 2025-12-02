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

MsgBox("Update Successful")
