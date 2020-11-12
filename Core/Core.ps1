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
$Global:xmlzipdir
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

#JSON 경로예시
$global:jsondir
$global:jsondata

#USERCONFIG
$global:userconfigdir = "data\Userconfig.json"
$global:userconfig = Get-Content -Path $global:userconfigdir -Encoding UTF8 | ConvertFrom-Json

#파싱 데이터 저장소
$Global:parsingdatajsondir = "data\parsingdata.json"
$Global:parsingdata = Get-Content -Path $Global:parsingdatajsondir -Encoding UTF8 | ConvertFrom-Json





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
    $Global:xmlzipdir = Get-ChildItem -Path $filename -Filter 점검결과*.zip | % {$_.FullName}
    $Global:wmvdir = Get-ChildItem -Path $filename -Filter *.wmv | % {$_.FullName}

}

#xml 파싱 (백신, 날짜 갖고오기)
function xmlparse {
    $tmpdir = "tmp"

    mkdir -Path $tmpdir

    Expand-Archive $Global:xmlzipdir -DestinationPath $tmpdir

    $beforexmlpath = Get-ChildItem -Path $tmpdir -Filter *_Result_Before.xml | %{$_.FullName}
    $Global:xmldata = [xml](Get-Content $beforexmlpath)

    $Global:parsingdata = Get-Content -Path $Global:parsingdatajsondir | ConvertFrom-Json
    $Global:parsingdata.xml | % {$_.date=$Global:xmldata.'PC-Check'.START_TIME.Split()[0]}
    $Global:parsingdata.xml | % {$_.vaccine=$Global:xmldata.'PC-Check'.Vaccine}
    $Global:parsingdata | ConvertTo-Json -Depth 32 | Set-Content $Global:parsingdatajsondir -Encoding UTF8
    
    rmdir -Path $tmpdir -Force -Recurse
}

#docx 파싱
function docxparse {
    $tmpdir = "tmp"

    $Global:parsingdata = Get-Content -Path $Global:parsingdatajsondir | ConvertFrom-Json
    #밑으로는 계속 추가

    
}

#사용자 정보가 올바른지 확인
function checkconfig {
    if (!$global:userconfig.user.name) {
        #유저 세팅이 안되어있음
        #메세지 박스
    }
}

