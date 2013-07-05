function Export-Xlsx {
    <#
        .SYNOPSIS
            Exports data to an Excel workbook

        .DESCRIPTION
            Exports data to an Excel workbook and applies cosmetics. Optionally add a title, autofilter, autofit and a chart.
            Allows for export to .xls and .xlsx format. If .xlsx is specified but not available (Excel 2003) the data will be exported to .xls.

        .PARAMETER InputData
            The data to be exported to Excel

        .PARAMETER Path
            The path of the Excel file. Defaults to %HomeDrive%\Export.xlsx.

        .PARAMETER WorksheetName
            The name of the worksheet. Defaults to filename in $Path without extension.

        .PARAMETER ChartType
            Name of an Excel chart to be added.

  	.PARAMETER TableStyle
			Apply a style to the created table. Does not work in Excel 2003.
			For an overview of styles see http://msdn.microsoft.com/fr-fr/library/documentformat.openxml.spreadsheet.tablestyle.aspx
			The Pivot styles are not used in this function.
			Overrides paramaters Border and HeaderColor.

        .PARAMETER Title
            Adds a title to the worksheet.

        .PARAMETER SheetPosition
            Adds the worksheet either to the 'begin' or 'end' of the Excel file.
            This parameter is ignored when creating a new Excel file.

        .PARAMETER TransposeColumnProperty
            Selects a property from the input object of which the value will be used as column title.
            Only works when using the Transpose parameter.

        .PARAMETER ChartOnNewSheet
            Adds a chart to a new worksheet instead of to the worksheet containing data.
            The Chart will be placed after the sheet containing data.
            Only works when parameter ChartType is used.

        .PARAMETER AppendWorksheet
            Appends a worksheet to an existing Excel file.
            This parameter is ignored when creating a new Excel file.

        .PARAMETER Borders
            Adds borders to all cells. Defaults to True.

        .PARAMETER HeaderColor
            Applies background color to the header row. Defaults to True.

        .PARAMETER AutoFit
            Apply autofit to columns. Defaults to True.

        .PARAMETER AutoFilter
            Apply autofilter. Defaults to True.

        .PARAMETER PassThrough
            When enabled returns file object of the generated file.

        .PARAMETER Force
            Overwrites existing Excel sheet. When this switch is not used but the Excel file already exists, a new file with datestamp will be generated.
            This switch is ignored when using the AppendWorksheet switch.

        .PARAMETER Transpose
            Transposes the data in Excel. This will automatically disable AutoFilter and HeaderColor

        .PARAMETER FreezePanes
            Freezes the title row.

        .EXAMPLE
            Get-Process | Export-Xlsx D:\Data\ProcessList.xlsx
            Exports a list of running processes to Excel

        .EXAMPLE
            Get-ADuser -Filter {enabled -ne $True} | Select-Object Name,Surname,GivenName,DistinguishedName | Export-Xlsx -Path 'D:\Data\Disabled Users.xlsx' -Title 'Disabled users of Contoso.com'
            Export all disabled AD users to Excel with optional title

        .EXAMPLE
            Get-Process | Sort-Object CPU -Descending | Export-Xlsx -Path D:\Data\Processes_by_CPU.xlsx
            Export a sorted processlist to Excel

        .EXAMPLE
            Export-Xlsx (Get-Process) -AutoFilter:$False -PassThrough | Invoke-Item
            Export a processlist to %HomeDrive%\Export.xlsx with AutoFilter disabled, and open the Excel file

        .NOTES
            Author : Gilbert van Griensven
            Website : http://www.itpilgrims.com/2013/01/export-xlsx-extended/
            Based on http://www.lucd.info/2010/05/29/beyond-export-csv-export-xls/
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$True)]
        [ValidateNotNullOrEmpty()]
        $InputData,
        [Parameter(Position=1)]
        [ValidateScript({
            $ReqExt = [System.IO.Path]::GetExtension($_)
            ($ReqExt -eq ".xls") -or
            ($ReqExt -eq ".xlsx")
        })]
        $Path=(Join-Path $env:HomeDrive "Export.xlsx"),
        [Parameter(Position=2)] $WorksheetName = [System.IO.Path]::GetFileNameWithoutExtension($Path),
        [Parameter(Position=3)]
        [ValidateSet("xl3DArea","xl3DAreaStacked","xl3DAreaStacked100","xl3DBarClustered",
            "xl3DBarStacked","xl3DBarStacked100","xl3DColumn","xl3DColumnClustered",
            "xl3DColumnStacked","xl3DColumnStacked100","xl3DLine","xl3DPie",
            "xl3DPieExploded","xlArea","xlAreaStacked","xlAreaStacked100",
            "xlBarClustered","xlBarOfPie","xlBarStacked","xlBarStacked100",
            "xlBubble","xlBubble3DEffect","xlColumnClustered","xlColumnStacked",
            "xlColumnStacked100","xlConeBarClustered","xlConeBarStacked","xlConeBarStacked100",
            "xlConeCol","xlConeColClustered","xlConeColStacked","xlConeColStacked100",
            "xlCylinderBarClustered","xlCylinderBarStacked","xlCylinderBarStacked100","xlCylinderCol",
            "xlCylinderColClustered","xlCylinderColStacked","xlCylinderColStacked100","xlDoughnut",
            "xlDoughnutExploded","xlLine","xlLineMarkers","xlLineMarkersStacked",
            "xlLineMarkersStacked100","xlLineStacked","xlLineStacked100","xlPie",
            "xlPieExploded","xlPieOfPie","xlPyramidBarClustered","xlPyramidBarStacked",
            "xlPyramidBarStacked100","xlPyramidCol","xlPyramidColClustered","xlPyramidColStacked",
            "xlPyramidColStacked100","xlRadar","xlRadarFilled","xlRadarMarkers",
            "xlStockHLC","xlStockOHLC","xlStockVHLC","xlStockVOHLC",
            "xlSurface","xlSurfaceTopView","xlSurfaceTopViewWireframe","xlSurfaceWireframe",
            "xlXYScatter","xlXYScatterLines","xlXYScatterLinesNoMarkers","xlXYScatterSmooth",
            "xlXYScatterSmoothNoMarkers")]
        [PSObject] $ChartType,
		[Parameter(Position=4)]
		[ValidateSet("TableStyleMedium28","TableStyleMedium27","TableStyleMedium26",
		"TableStyleMedium25","TableStyleMedium24","TableStyleMedium23",
		"TableStyleMedium22","TableStyleMedium21","TableStyleMedium20",
		"TableStyleMedium19","TableStyleMedium18","TableStyleMedium17",
		"TableStyleMedium16","TableStyleMedium15","TableStyleMedium14",
		"TableStyleMedium13","TableStyleMedium12","TableStyleMedium11",
		"TableStyleMedium10","TableStyleMedium9","TableStyleMedium8",
		"TableStyleMedium7","TableStyleMedium6","TableStyleMedium5",
		"TableStyleMedium4","TableStyleMedium3","TableStyleMedium2",
		"TableStyleMedium1","TableStyleLight21","TableStyleLight20",
		"TableStyleLight19","TableStyleLight18","TableStyleLight17",
		"TableStyleLight16","TableStyleLight15","TableStyleLight14",
		"TableStyleLight13","TableStyleLight12","TableStyleLight11",
		"TableStyleLight10","TableStyleLight9","TableStyleLight8",
		"TableStyleLight7","TableStyleLight6","TableStyleLight5",
		"TableStyleLight4","TableStyleLight3","TableStyleLight2",
		"TableStyleLight1","TableStyleDark11","TableStyleDark10",
		"TableStyleDark9","TableStyleDark8","TableStyleDark7",
		"TableStyleDark6","TableStyleDark5","TableStyleDark4",
		"TableStyleDark3","TableStyleDark2","TableStyleDark1")]
		[String] $TableStyle,
        [Parameter(Position=5)] $Title,
        [Parameter(Position=6)] [ValidateSet("begin","end")] $SheetPosition="begin",
        [Parameter(Position=7)] [String] $TransposeColumnProperty,
        [Switch] $ChartOnNewSheet,
        [Switch] $AppendWorksheet,
        [Switch] $Borders=$True,
        [Switch] $HeaderColor=$True,
        [Switch] $AutoFit=$True,
        [Switch] $AutoFilter=$True,
        [Switch] $PassThrough,
        [Switch] $Force,
        [Switch] $Transpose,
        [Switch] $FreezePanes
    )
    Begin {
        Add-Type -AssemblyName Microsoft.Office.Interop.Excel 
        Function Convert-NumberToA1 {
          Param([parameter(Mandatory=$true)] [int]$number)
          $a1Value = $null
          While ($number -gt 0) {
            $multiplier = [int][system.math]::Floor(($number / 26))
            $charNumber = $number - ($multiplier * 26)
            If ($charNumber -eq 0) { $multiplier-- ; $charNumber = 26 }
            $a1Value = [char]($charNumber + 96) + $a1Value
            $number = $multiplier
          }
          Return $a1Value
        }

        $Script:WorkingData = @()
    }
    Process {
        $Script:WorkingData += $InputData
    }
    End {
        $Props = $Script:WorkingData[0].PSObject.properties | % { $_.Name }
        If ($Transpose) {
            $Rows = $Props.Count
            $Cols = $Script:WorkingData.Count+1
            $AutoFilter = $False
            If (($TransposeColumnProperty) -and ($Props.Contains($TransposeColumnProperty))) {
                $Rows++
            } Else {
                $TransposeColumnProperty = $Null
                $HeaderColor = $False
            }
        } Else {
            $Rows = $Script:WorkingData.Count+1
            $Cols = $Props.Count
        }
        $A1Cols = Convert-NumberToA1 $Cols
        $Array = New-Object 'object[,]' $Rows,$Cols

        $Col = 0
        $Row = 0
        If (($Transpose) -and ($TransposeColumnProperty)) {
            $Row++
        }

        $Props | % {
            $Array[$Row,$Col] = $_.ToString()
            If ($Transpose) {
                $Row++
            } Else {
                $Col++
            }
        }

        $Row = 1
        $Script:WorkingData | % {
            $Item = $_
            $Col = 0
            $Props | % {
                If (($Transpose) -and ($TransposeColumnProperty) -and ($Col -eq 0)) {
                    $Array[$Col,$Row] = $Item.($TransposeColumnProperty).ToString()
                    $Col++
                }
                If ($Item.($_) -eq $Null) {
                    If ($Transpose) {
                        $Array[$Col,$Row] = ""
                    } Else {
                        $Array[$Row,$Col] = ""
                    }
                } Else {
                    If ($Transpose) {
                        $Array[$Col,$Row] = $Item.($_).ToString()
                    } Else {
                        $Array[$Row,$Col] = $Item.($_).ToString()
                    }
                }
                $Col++
            }
            $Row++
        }

        $ExcelApplication = New-Object -ComObject Excel.Application
        $ExcelApplication.DisplayAlerts = $False
        $ExcelApplicationFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlWorkbookNormal

        If ([System.IO.Path]::GetExtension($Path) -eq '.xlsx') {
            If ($ExcelApplication.Version -lt 12) {
                $Path = $Path.Replace(".xlsx",".xls")
            } Else {
                $ExcelApplicationFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlWorkbookDefault
            }
        }

        If (Test-Path -Path $Path -PathType "Leaf") {
            If ($AppendWorkSheet) {
                $Workbook = $ExcelApplication.Workbooks.Open($Path)
                If ($SheetPosition -eq "end") {
                    $Workbook.Worksheets.Add([System.Reflection.Missing]::Value,$Workbook.Sheets.Item($Workbook.Sheets.Count)) | Out-Null
                } Else {
                    $Workbook.Worksheets.Add($Workbook.Worksheets.Item(1)) | Out-Null
                }
            } Else {
                If (!($Force)) {
                    $Path = $Path.Insert($Path.LastIndexOf(".")," - $(Get-Date -Format "ddMMyyyy-HHmm")")
                }
                $Workbook = $ExcelApplication.Workbooks.Add()
                While ($Workbook.Worksheets.Count -gt 1) { $Workbook.Worksheets.Item(1).Delete() }
            }
        } Else {
            $Workbook = $ExcelApplication.Workbooks.Add()
            While ($Workbook.Worksheets.Count -gt 1) { $Workbook.Worksheets.Item(1).Delete() }
        }

        $Worksheet = $Workbook.ActiveSheet
        Try { $Worksheet.Name = $WorksheetName }
        Catch { }

        If ($Title) {
            $Worksheet.Cells.Item(1,1) = $Title
            $TitleRange = $Worksheet.Range("a1","$($A1Cols)2")
            $TitleRange.Font.Size = 18
            $TitleRange.Font.Bold=$True
            $TitleRange.Font.Name = "Cambria"
            $TitleRange.Font.ThemeFont = 1
            $TitleRange.Font.ThemeColor = 4
            $TitleRange.Font.ColorIndex = 55
            $TitleRange.Font.Color = 8210719
            $TitleRange.Merge()
            $TitleRange.VerticalAlignment = -4160
			While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($TitleRange)) {}
            $UsedRange = $Worksheet.Range("a3","$($A1Cols)$($Rows + 2)")
            If ($HeaderColor) {
                $Worksheet.Range("a3","$($A1Cols)3").Interior.ColorIndex = 48
                $Worksheet.Range("a3","$($A1Cols)3").Font.Bold = $True
            }
            If (($FreezePanes) -and ((($Transpose) -and ($TransposeColumnProperty)) -or (!($Transpose)))) {
                $Worksheet.Range("a4:a4").Select() | Out-Null
                $ExcelApplication.ActiveWindow.FreezePanes = $True
            }
        } Else {
            $UsedRange = $Worksheet.Range("a1","$($A1Cols)$($Rows)")
            If ($HeaderColor) {
                $Worksheet.Range("a1","$($A1Cols)1").Interior.ColorIndex = 48
                $Worksheet.Range("a1","$($A1Cols)1").Font.Bold = $True
            }
            If (($FreezePanes) -and ((($Transpose) -and ($TransposeColumnProperty)) -or (!($Transpose)))) {
                $Worksheet.Range("a2:a2").Select() | Out-Null
                $ExcelApplication.ActiveWindow.FreezePanes = $True
            }
        }

        $UsedRange.Value2 = $Array
        $UsedRange.HorizontalAlignment = -4131

        If ($Borders) {
            $UsedRange.Borders.LineStyle = 1
            $UsedRange.Borders.Weight = 2
        }

        If ($AutoFilter) { $UsedRange.AutoFilter() | Out-Null }

        If ($AutoFit) { $Worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null }

		If (($TableStyle) -and ($ExcelApplication.Version -ge 12)) {
			$ListObject = $Worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $UsedRange, $Null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $Null)
			$ListObject.TableStyle = $TableStyle
			While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($ListObject)) {}
		}

        If ($ChartType) {
            [Microsoft.Office.Interop.Excel.XlChartType]$ChartType = $ChartType
            If ($ChartOnNewSheet) {
                $Workbook.Charts.Add().ChartType = $ChartType
                $Workbook.ActiveChart.setSourceData($UsedRange)
                Try { $Workbook.ActiveChart.Name = "$($WorksheetName) - Chart" }
                Catch { }
                $Workbook.ActiveChart.Move([System.Reflection.Missing]::Value,$Workbook.Sheets.Item($Worksheet.Name))
            } Else {
                $Worksheet.Shapes.AddChart($ChartType).Chart.setSourceData($UsedRange) | Out-Null
            }
        }

        $Workbook.SaveAs($Path,$ExcelApplicationFixedFormat)
        $Workbook.Close()
        $ExcelApplication.Quit()

        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($UsedRange)) {}
        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheet)) {}
        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook)) {}
        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApplication)) {}
        [GC]::Collect()

        If ($PassThrough) { Return Get-Item $Path }
    }
}
