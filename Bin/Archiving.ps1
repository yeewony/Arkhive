#Addtype
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()

#파웨쉘 숨기기
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

#변수----------변수----------변수----------변수----------변수----------변수----------변수

$Global:iotpdfdir
$Global:iotzipdir

$Global:docxdir

$Global:xmlzipdir

$Global:wmvdir

$Global:localdir = Get-Location
#

#Script----------Script----------Script----------Script----------Script----------Script

function getfiles {
    #파일 경로 갖고오기
    $Global:IoTzipdir = Get-ChildItem -Path $filename -Filter data.zip | % {$_.FullName}
    $Global:IoTpdfdir = Get-ChildItem -Path $filename -Filter *.pdf | % {$_.FullName}
    $Global:docxdir = Get-ChildItem -Path $filename -Filter *.docx | % {$_.FullName}
    $Global:xmlzipdir = Get-ChildItem -Path $filename -Filter 점검결과*.zip | % {$_.FullName}
    $Global:wmvdir = Get-ChildItem -Path $filename -Filter *.wmv | % {$_.FullName}  

}
