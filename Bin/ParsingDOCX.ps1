#Add-Type
Add-Type -Path D:\archivenindex\Bin\DocumentFormat.OpenXml.dll


#파워쉘 숨기기
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

#변수----------변수----------변수----------변수----------변수----------변수----------변수

$Global:localpath = Get-Location

$Global:DOCXDATA_DATE = $null
$Global:DOCXDATA_DATE_YEAR = $null
$Global:DOCXDATA_DATE_MONTH = $null
$Global:DOCXDATA_DATE_DAY = $null
$Global:DOCXDATA_DATE_HOUR = $null
$Global:DOCXDATA_VACCINE = $null
$Global:DOCXDATA_RESULT = $null

$Global:DOCXDATA_CELLNUM = 0

#Script----------Script----------Script----------Script----------Script----------Script

function parsingdocx {

    $docxpath = Get-ChildItem -Path $Global:localpath -Filter *.docx | % {$_.FullName}

    [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]$docxdata = $null

    $docxdata = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($docxpath, $true)
    
    $Global:DOCXDATA_DATE = $docxdata.MainDocumentPart.Document.Body.Elements()[3].InnerText #시작시간
    $Global:DOCXDATA_VACCINE = $docxdata.MainDocumentPart.Document.Body.Elements()[6].InnerText #백신
    $Global:DOCXDATA_RESULT = $docxdata.MainDocumentPart.Document.Body.Elements()[8] #결과
    
    if ($Global:DOCXDATA_RESULT.InnerText -eq ""){

        $Global:DOCXDATA_RESULT = $docxdata.MainDocumentPart.Document.Body.Elements()[7]
        $Global:DOCXDATA_CELLNUM = 1
    }

    $docxdata.Close()
    $docxdata = $null
    
}

function savedate {
    $DATE_YEAR = $Global:DOCXDATA_DATE.Split()[4].Split("-")[0]
    $DATE_MONTH = $Global:DOCXDATA_DATE.Split()[4].Split("-")[1]
    $DATE_DAY = $Global:DOCXDATA_DATE.Split()[4].Split("-")[2]
    $DATE_HOUR = $Global:DOCXDATA_DATE.Split()[5].Split(":")[0]
    $VACCINE = $Global:DOCXDATA_VACCINE.Split()[-1]

    

}

parsingdocx
savedate

