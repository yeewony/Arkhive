#Addtype
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()

#DLL
Add-Type -Path DocumentFormat.OpenXml.dll

#PS Hidden
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)



#변수----------변수----------변수----------변수----------변수----------변수----------변수
#IoT(data,pdf), Docx, XML, WMV

#iot
$Global:IoTzipchk
$Global:IoTpdfchk
$Global:IoTzipfilechk
$Global:IoTpdffilechk
$Global:IoTzipdir
$Global:IoTpdfdir

#docx
$Global:docxchk
$Global:docxfilechk
$Global:docxdir
$Global:docxdata
$Global:docxcellnum

#xml
$Global:xmlchk
$Global:xmlfilechk
$Global:xmldir
$Global:xmldata

#wmv
$Global:wmvchk
$Global:wmvfilechk
$Global:wmvdir

#기타 변수들(날짜, 주차, 여러가지)
#월요일에 대한 주차 반영이 안됨, gmt 문제로 추정됨, 해당 코드는 수정이 되야함.
$Global:weekofyear = Get-Date -UFormat %V
$Global:Date


#JSON----------JSON----------JSON---------JSON----------JSON----------JSON----------JSON

#JSON 경로
$global:jsondir
$global:jsondata


#Script----------Script----------Script----------Script----------Script----------Script

#파일 경로 갖고오기
function getfilesdir {
    #파일 경로 전체 가져오기
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) {
        }
    }

    #파일 로딩 체크
    #ui에 파일 로딩 체크 한것을 표시하게 하기

    #경로 삽입
    $Global:IoTzipdir = Get-ChildItem -Path $filename -Filter data.zip | % {$_.FullName}
    $Global:IoTpdfdir = Get-ChildItem -Path $filename -Filter *.pdf | % {$_.FullName}
    $Global:docxdir = Get-ChildItem -Path $filename -Filter *.docx | % {$_.FullName}
    $Global:xmldir = Get-ChildItem -Path $filename -Filter *_Result_Before.xml | % {$_.FullName}
    $Global:wmvdir = Get-ChildItem -Path $filename -Filter *.wmv | % {$_.FullName}

}

#xml 파싱 (백신, 날짜 갖고오기)
function xmlparse {

}

#docx 파싱
