#Addtype
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()

#파웨쉘 숨기기
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

#변수----------변수----------변수----------변수----------변수----------변수----------변수

$Global:iotpdfdir = $null
$Global:iotzipdir = $null
$Global:docxdir = $null
$Global:xmlzipdir = $null
$Global:wmvdir = $null

$Global:localdir = Get-Location

$Global:indexjsonpath = "D:\archivenindex\Bin\Data\index.json"

$Global:arkdir = "D:\Ark"

#Script----------Script----------Script----------Script----------Script----------Script

function getfiles {
    #파일 경로 갖고오기
    $Global:IoTzipdir = Get-ChildItem -Path $Global:localdir -Filter data.zip | % {$_.FullName}
    $Global:IoTpdfdir = Get-ChildItem -Path $Global:localdir -Filter *.pdf | % {$_.FullName}
    $Global:docxdir = Get-ChildItem -Path $Global:localdir -Filter *.docx | % {$_.FullName}
    $Global:xmlzipdir = Get-ChildItem -Path $Global:localdir -Filter 점검결과*.zip | % {$_.FullName}
    $Global:wmvdir = Get-ChildItem -Path $Global:localdir -Filter *.wmv | % {$_.FullName}  

}

function copyfiles {
    Copy-Item -Path D:\archivenindex\bin\Data\index.json -Destination $Global:localdir -Force -Recurse
    Copy-Item -Path D:\archivenindex\bin\ParsingXML.ps1 -Destination $Global:localdir -Force -Recurse
}

function copytoark {
    if(![System.IO.File]::Exists($Global:arkdir)){
        New-Item -ItemType Directory -Path $Global:arkdir
    }

    Copy-Item -Path $Global:iotpdfdir,$Global:iotzipdir,$Global:docxdir,$Global:xmlzipdir,$Global:wmvdir -Destination $Global:arkdir -Recurse -Force

}


getfiles
copyfiles
./ParsingXML.ps1
#docx파싱
#iot파싱??
#복사했던 파일 정리
#index파일 정리
copytoark