#Addtype
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()

#파웨쉘 숨기기
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

#변수----------변수----------변수----------변수----------변수----------변수----------변수

$Global:localdir = Get-Location
$Global:indexjson = "index.json"
$Global:XMLDATA = $null

#Script----------Script----------Script----------Script----------Script----------Script

function parsingxml {
    $getzip = Get-ChildItem -Path $Global:localdir -Filter 점검결과*.zip | % {$_.FullName}
    $tmpdir = "tmp"

    mkdir -Path $tmpdir

    Expand-Archive $getzip -DestinationPath $tmpdir

    $beforexmlpath = Get-ChildItem -Path $tmpdir -Filter *_Result_Before.xml | % {$_.FullName}
    $Global:XMLDATA = [xml](Get-Content $beforexmlpath)    

    rmdir -Path $tmpdir -Force -Recurse

}

function savexmldata {
    $index = Get-Content -Path $Global:indexjson -Encoding UTF8 | ConvertFrom-Json
    $index.data | % {$_.date=$Global:XMLDATA.'PC-Check'.START_TIME.Split()[0]}
    $index.data | % {$_.vaccine=$Global:XMLDATA.'PC-Check'.Vaccine}
    $index.indexinfo.date | % {$_.all=$Global:XMLDATA.'PC-Check'.START_TIME.Split()[0]}

    $datetime = $Global:XMLDATA.'PC-Check'.START_TIME.Split()[0]
    $datehour = $Global:XMLDATA.'PC-Check'.START_TIME.Split()[1].Split(":")[0]
    $index.indexinfo.date | % {$_.year=$datetime.Split("-")[0]}
    $index.indexinfo.date | % {$_.month=$datetime.Split("-")[1]}
    $index.indexinfo.date | % {$_.day=$datetime.Split("-")[2]}
    $index.indexinfo.date | % {$_.hour=$datehour+":00"}
    $index | ConvertTo-Json -Depth 32 | Set-Content $Global:indexjson -Encoding UTF8
}

parsingxml
savexmldata