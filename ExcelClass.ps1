Class ExcelReporting
{
    $Excel

    ExcelReporting( [string]$Name)
    {
        $this.SetCulture()
        $this.Excel = New-Object -ComObject excel.application

        $this.Excel.Visible = $True

        #Add a workbook
        $this.Excel.SheetsInNewWorkbook = 1
        $workbook = $this.Excel.Workbooks.Add()

        #Connect to first worksheet to rename and make active
        $serverInfoSheet = $workbook.Worksheets.Item(1)
        $serverInfoSheet.Name = $Name

    }
    [void]SetFreezepane()
    {
        $worksheet = $this.Excel.ActiveSheet
        $worksheet.Activate() | Out-Null
        $worksheet.application.activewindow.splitcolumn = 0
        $worksheet.application.activewindow.splitrow = 1
        $worksheet.application.activewindow.freezepanes = $true
    }

    [void]SetCulture()
    {
        #$curculture = Get-Culture

        $culture = [System.Globalization.CultureInfo]::GetCultureInfo(1033)
        [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture
        [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    }

    [void]SetHeaders( [string[]]$Headers )
    {
        $row = 1
        $column = 1

        $worksheet = $this.Excel.ActiveSheet
        foreach( $header in $Headers )
        {
            $worksheet.Cells.Item($row,$column)= $header
            $worksheet.Cells.Item($row,$column).Interior.ColorIndex = 48
            $worksheet.Cells.Item($row,$column).Font.Bold=$True
            $Column++
        }
    }

    [void]SetItem( [int]$row, [int]$column, [string]$cellContent )
    {
        $worksheet = $this.Excel.ActiveSheet

        $worksheet.Cells.Item($row,$column) = $cellContent
    }

    [void]AutoFit()
    {
        $this.Excel.ActiveSheet.UsedRange.Columns.AutoFit()
    }
}

# Voorbeeld waarbij de actieve processen in Excel getoond worden.
$processen = Get-Process | Select-Object Name, CPU

[string[]]$headers = 'Naam','CPU tijd'

$processenRpt = [ExcelReporting]::new( 'Active processen')
$processenRpt.SetFreezePane()
$processenRpt.SetHeaders( $headers )

$row = 2
$column = 1

foreach( $rptItemitem in $processen )
{
    $processenRpt.SetItem( $row, $column++, $rptItemitem.Name )
    $processenRpt.SetItem( $row, $column++, $rptItemitem.CPU )
    
    $column = 1
    $row += 1
}

$processenRpt.AutoFit()
