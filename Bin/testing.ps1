#test

$local = Get-Location
$iotzipdir = $null
function get {    
$iotzipdir = Get-ChildItem -Path $local -Filter data.zip | % {$_.FullName}
}

get

Write-Host $local
Write-Host $iotzipdir