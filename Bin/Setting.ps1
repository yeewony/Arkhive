#Addtype
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -AssemblyName PresentationFramework
[System.Windows.Forms.Application]::EnableVisualStyles()

#파웨쉘 숨기기
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

#변수----------변수----------변수----------변수----------변수----------변수----------변수

$Global:indexjsonpath = "data\index.json"
$Global:userconfigjsonpath = "data\Userconfig.json"


#Script----------Script----------Script----------Script----------Script----------Script

function importuserdatatoindexjson {
    $userconfig = Get-Content -Path $Global:userconfigjsonpath -Encoding UTF8 | ConvertFrom-Json
    $index = Get-Content -Path $Global:indexjsonpath -Encoding UTF8 | ConvertFrom-Json
    $index.indexinfo.inspector | % {$_.name=$userconfig.inspector.name}
    $index.indexinfo.inspector | % {$_.id=$userconfig.inspector.id}
    $index.indexinfo.inspector | % {$_.level=$userconfig.inspector.level}
    $index | ConvertTo-Json -Depth 32 | Set-Content $Global:indexjsonpath -Encoding UTF8
}