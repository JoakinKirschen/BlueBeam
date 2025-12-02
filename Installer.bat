powershell -Command "Invoke-WebRequest 'https://github.com/JoakinKirschen/BlueBeam/blob/master/Package.exe' -OutFile %userprofile%\Downloads\SSD_zippertool.exe"
start /d %userprofile%\Downloads\SSD_zippertool.exe
xcopy /s %userprofile%\Downloads\ %appdata%\Roaming\Bluebeam Software\Revu\21\Templates





xcopy /s %userprofile%\Downloads\ %appdata%\Roaming\Bluebeam Software\Revu\21

