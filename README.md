# ExcelClass
Reporting using Excel


Convert an array to Excel

```
$ServicesRpt = [ExcelReporting]::new( 'Services')
$ServicesRpt.SetFreezePane()
$ServicesRpt.FromArray( (Get-Service ))
$ServicesRpt.AutoFit()
```