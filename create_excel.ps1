$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.ActiveSheet
$worksheet.Name = "투입인원 명단"

$worksheet.Range("A1:F1").Merge()
$worksheet.Range("A1").Value2 = "가산~가평 천연가스 공급시설 건설공사 비파괴검사 기술용역 투입인원 명단"
$worksheet.Range("A1").Font.Size = 16
$worksheet.Range("A1").Font.Bold = $true
$worksheet.Range("A1").HorizontalAlignment = -4108
$worksheet.Range("A1").VerticalAlignment = -4108
$worksheet.Rows.Item(1).RowHeight = 40

$headers = @("연번", "직책 (담당분야)", "성명", "생년월일", "서명 (인)", "비고")
for ($i=0; $i -lt $headers.Length; $i++) {
    $cell = $worksheet.Cells.Item(3, $i+1)
    $cell.Value2 = $headers[$i]
    $cell.Font.Bold = $true
    $cell.HorizontalAlignment = -4108
    $cell.VerticalAlignment = -4108
}

$worksheet.Columns.Item(1).ColumnWidth = 8
$worksheet.Columns.Item(2).ColumnWidth = 25
$worksheet.Columns.Item(3).ColumnWidth = 15
$worksheet.Columns.Item(4).ColumnWidth = 20
$worksheet.Columns.Item(5).ColumnWidth = 20
$worksheet.Columns.Item(6).ColumnWidth = 25

$data = @(
    @("1", "총괄 책임자", "", "", "", ""),
    @("2", "방사선투과검사(RT)", "", "", "", ""),
    @("3", "초음파탐상검사(UT)", "", "", "", ""),
    @("4", "자기탐상검사(MT)", "", "", "", ""),
    @("5", "침투탐상검사(PT)", "", "", "", "")
)

for ($i=0; $i -lt 10; $i++) {
    if ($i -lt $data.Length) {
        $row = $data[$i]
    } else {
        $row = @( ($i+1).ToString(), "", "", "", "", "" )
    }
    
    for ($j=0; $j -lt 6; $j++) {
        $cell = $worksheet.Cells.Item($i+4, $j+1)
        $cell.Value2 = $row[$j]
        $cell.HorizontalAlignment = -4108
        $cell.VerticalAlignment = -4108
    }
}

$range = $worksheet.Range("A3:F13")
$borders = @(7, 8, 9, 10, 11, 12)
foreach ($borderId in $borders) {
    $border = $range.Borders.Item($borderId)
    $border.LineStyle = 1
    $border.Weight = 2
}

for ($r=3; $r -le 13; $r++) {
    $worksheet.Rows.Item($r).RowHeight = 30
}

$savePath = "c:\Users\-\OneDrive\바탕 화면\home\투입인원_명단.xlsx"
if (Test-Path $savePath) {
    Remove-Item $savePath
}
$workbook.SaveAs($savePath)
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Output "SUCCESS: $savePath"
