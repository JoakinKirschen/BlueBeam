powershell -Command "Invoke-WebRequest 'https://swecogroup.sharepoint.com/sites/Gr_Gr_Structural_Buildings/Knowledge/_layouts/15/download.aspx?SourceUrl=%2Fsites%2FGr%5FGr%5FStructural%5FBuildings%2FKnowledge%2FKnowledge%20Structural%2FStructural%20System%20Drawing%20%2D%20toolchest%2FSSD%5Fzippertool%2Eexe' -OutFile %userprofile%\Downloads\SSD_zippertool.exe"
start /d %userprofile%\Downloads\SSD_zippertool.exe
xcopy /s %userprofile%\Downloads\ %appdata%\Roaming\Bluebeam Software\Revu\21\Templates





xcopy /s %userprofile%\Downloads\ %appdata%\Roaming\Bluebeam Software\Revu\21
