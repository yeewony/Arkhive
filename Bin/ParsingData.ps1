#Add-Type


#파워쉘 숨기기
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

#변수----------변수----------변수----------변수----------변수----------변수----------변수

$Global:localdir = Get-Location
$Global:indexjson = "index.json"
$Global:DOCXDATA = $null

$global:localappdata = $env:LOCALAPPDATA

#Script----------Script----------Script----------Script----------Script----------Script

function parsingdocx {
    
}