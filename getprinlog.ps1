#set start day
$sdate =(Get-Date).AddDays(-7)

#Get print eventlog
$logs = Get-WinEvent -FilterHashtable @{logname="Microsoft-Windows-PrintService/Operational";id=307; starttime=$sdate}
$printlog = @()
foreach ($log in $logs){
 
$obj = New-Object PSObject -Property @{UserName=""; PCName=""; PageCount=""; Time=""; PrinterName=""; Message=""}
    $obj.UserName = $log.Properties[2].Value
    $obj.PCName = $log.Properties[3].Value
    $obj.PageCount = $log.Properties[7].Value
    $obj.Time = $log.TimeCreated
    $obj.PrinterName = $log.Properties[4].Value
    $obj.Message = $log.Message
    $printlog += $obj | Select UserName, PCName, Pagecount, Time, printername
    }

$date = get-date -Format "dd.MM.yyyy"
$printlog | Export-Csv "C:\PrintLogTemp\$date.csv" -NoTypeInformation

try { 
#Export Csv to Excel
#Create new Excel Object
$excel = New-Object -ComObject excel.application
$excel.visible = $false

#create new workbook
$workbook = $excel.workbooks.add()

#new sheet
$worksheet = $workbook.worksheets.Item(1)

#rename sheet
$worksheet.name = 'Temp-PrintLog'

#Add name to column 
        $worksheet.cells.item(1,1) = "UserName"
        $worksheet.cells.item(1,2) = "PCName"
        $worksheet.cells.item(1,3) = "PageCount"
        $worksheet.cells.item(1,4) = "Time"
        $worksheet.cells.item(1,5) = "PrinterName"

#make font italic and bold
$worksheet.rows.item(1).font.italic = $true
$worksheet.rows.item(1).font.bold = $true

#autosize column
$form = $worksheet.usedRange
$form.entirecolumn.autofit() | Out-Null

$proc = Import-csv -Path "C:\PrintLogTemp\$date.csv"
    $s = 2
        foreach($procc in $proc)
                {
                $worksheet.cells.item($s,1) = $procc.UserName
                $worksheet.cells.item($s,2) = $procc.PCName
                $worksheet.cells.item($s,3) = $procc.PageCount
                $worksheet.cells.item($s,4) = $procc.Time
                $worksheet.cells.item($s,5) = $procc.PrinterName
                $s++
                $formm = $worksheet.usedRange
                $formm.entirecolumn.autofit() | Out-Null
                }
    

$workbook.saveas("C:\Temp Print Logs\$date.xlsx")
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel
} 
        catch {
        "error in last block" | Add-Content c:\logs.txt }
spps -n excel