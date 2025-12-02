Add-Type -AssemblyName System.IO.Compression.FileSystem

Write-Output "Start downloading"

$Url = "https://github.com/JoakinKirschen/BlueBeam/archive/master.zip"
$filename = Split-Path $Url -Leaf

$TempPath = [System.IO.Path]::GetTempPath()
$myPath = Join-Path $TempPath $filename

Invoke-WebRequest -Uri $Url -OutFile $myPath

Write-Output "Start extracting"
Write-Output "Extracting new version"

$ExtractTo = $TempPath
$ExtractFolder = Join-Path $ExtractTo "BlueBeam-master"

if (Test-Path $ExtractFolder) {
    Remove-Item -Path $ExtractFolder -Recurse -Force
}

# Extract the zip file
[System.IO.Compression.ZipFile]::ExtractToDirectory($myPath, $ExtractTo)

$CopySource = Join-Path $ExtractFolder ""
$AppData = [Environment]::GetFolderPath("ApplicationData")

Write-Output "Copying new version"
Write-Output $AppData
Write-Output $ExtractFolder

$destfoldertemplate = Join-Path $AppData "Bluebeam Software\Revu\21\Templates\"
$destfolderprofile = Join-Path $AppData "Bluebeam Software\Revu\21\"

function sdcopy {
    param (
        [string]$filename,
        [string]$destfolder
    )
    $source = Join-Path $CopySource $filename
    $dest = Join-Path $destfolder $filename

    if (Test-Path $dest) {
        $file = Get-Item $dest
        if (-not ($file.Attributes -band [System.IO.FileAttributes]::ReadOnly)) {
            Copy-Item -Path $source -Destination $dest -Force
        }
        else {
            # Remove ReadOnly attribute
            $file.Attributes = $file.Attributes -bxor [System.IO.FileAttributes]::ReadOnly
            Copy-Item -Path $source -Destination $dest -Force
            # Reapply ReadOnly attribute
            $file.Attributes = $file.Attributes -bor [System.IO.FileAttributes]::ReadOnly
        }
    }
    else {
        Copy-Item -Path $source -Destination $dest -Force
    }
}

sdcopy "_SSD - A0 landscape.pdf" $destfoldertemplate
sdcopy "_SSD - A0-26 landscape.pdf" $destfoldertemplate
sdcopy "_SSD - A3 landscape.pdf" $destfoldertemplate
sdcopy "_SSD - A4 portrait.pdf" $destfoldertemplate
sdcopy "_SSD Bluebeam toolchest building template 20251126.pdf" $destfoldertemplate
sdcopy "_SSD stamp NOK.pdf" $destfoldertemplate
sdcopy "_SSD stamp OK, with comments.pdf" $destfoldertemplate
sdcopy "_SSD stamp OK.pdf" $destfoldertemplate

sdcopy "Sweco STRU profile.bpx" $destfolderprofile

[System.Windows.MessageBox]::Show("Update Successful")
