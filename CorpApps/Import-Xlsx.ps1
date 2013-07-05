function Import-Xlsx {
    <#
        .SYNOPSIS
            Import data from an Excel worksheet to a custom PSObject.

        .DESCRIPTION
            Imports data from an Excel worksheet to a custom PSObject. Imports from both .xls and .xlsx formats.

        .PARAMETER Path
            The path of the Excel file.
            If the parameters SheetName and SheetIndex are omitted you will be queried for which sheet to import. If there is only one worksheet in the workbook it will import that sheet without querying.

        .PARAMETER SheetName
            The name of the worksheet to import. Will throw an error if the worksheet name can not be found, or if the sheet with that name is not a worksheet.
            Can not be used in conjuction with parameter SheetIndex.

        .PARAMETER SheetIndex
            The index of the worksheet to import. Will throw an error if the worksheet index can not be found, or if the sheet with that index is not a worksheet.
            Can not be used in conjunction with parameter SheetName

        .PARAMETER IsTransposed
            Use this switch to indicate that the worksheet contains a transposed table.
            See Export-Xlsx.

        .PARAMETER HasTransposeColumnProperty
            Use this switch to indicate that the transposed table includes a header for each column. Enabled by default.
            This switch will be ignored when the parameter IsTransposed is not used.
            See Export-Xlsx.

        .PARAMETER HasTitle
            Use this switch to indicate that the worksheet contains a title.
            See Export-Xlsx.

        .EXAMPLE
            Import-Xlsx D:\UserData.xlsx
            Imports data from a .xlsx file. Will query for worksheet if there is more than one valid worksheet in the Excel file.

        .EXAMPLE
            Import-Xlsx D:\UserData.xlsx -SheetName "Disabled Users"
            Imports data from the worksheet with the name "Disabled Users"

        .NOTES
            Author : Gilbert van Griensven
            Website : http://www.itpilgrims.com/2013/01/import-xlsx/
    #>
    [CmdletBinding(DefaultParametersetName="Default")]
    Param (
        [Parameter(Position=0,Mandatory=$True)]
        [ValidateScript({
            $ReqExt = [System.IO.Path]::GetExtension($_)
            ($ReqExt -eq ".xls") -or
            ($ReqExt -eq ".xlsx")
        })]
        $Path,
        [Parameter(ParameterSetName="ByName")]
        [String] $SheetName,
        [Parameter(ParameterSetName="ByIndex")]
        [Int] $SheetIndex,
        [Switch] $IsTransposed,
        [Switch] $HasTransposeColumnProperty=$True,
        [Switch] $HasTitle
    )
    Function ReadData ($FromFile) {
        Add-Type -AssemblyName Microsoft.Office.Interop.Excel 
        $ExcelApplication = New-Object -ComObject Excel.Application
        $ExcelApplication.DisplayAlerts = $False
        $Workbook = $ExcelApplication.Workbooks.Open($FromFile)
        $Worksheets = $Workbook.Worksheets
        If ($Worksheets.Count -ge 1) {
            Switch ($PsCmdlet.ParameterSetName) {
                "Default" {
                    If ($Worksheets.Count -gt 1) {
                        $ChoiceDesc = New-Object System.Collections.ObjectModel.Collection[System.Management.Automation.Host.ChoiceDescription]
                        $SheetChoice = 0
                        $Script:Sheets = @()
                        $Worksheets |
                        % {
                            $Sheet = $_
                            $Script:Sheets += New-Object PSObject -Property @{Choice=$SheetChoice;SheetIndex=$Sheet.Index}
                            $ChoiceDesc.Add((New-Object "System.Management.Automation.Host.ChoiceDescription" -ArgumentList "&$($SheetChoice) $($Sheet.Name)"))
                            $SheetChoice++
                        }
                        $Result = $Host.UI.PromptForChoice("Import data from $($FromFile)","Please select the sheet to import:",$ChoiceDesc,0)
                        $SelectedSheet = $Script:Sheets | Where-Object -FilterScript {$_.Choice -eq $Result} | Select-Object -ExpandProperty SheetIndex
                        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Sheet)) {}
                    } Else {
                        $SelectedSheet = $Worksheets | % {$_.Index}
                    }
                }
                "ByName" {
                    $SelectedSheet = ($Worksheets | Where-Object -FilterScript {$_.Name -eq $SheetName}).Index
                    If (!($SelectedSheet)) { $ExceptionMessage = "A worksheet with the name '$($SheetName)' can not be found in workbook $($Path) or is not of the type 'Worksheet'." }
                }
                "ByIndex" {
                    $SelectedSheet = ($Worksheets | Where-Object -FilterScript {$_.Index -eq $SheetIndex}).Index
                    If (!($SelectedSheet)) { $ExceptionMessage = "A worksheet with index '$($SheetIndex)' can not be found in workbook $($Path) or is not of the type 'Worksheet'." }
                }
            }
        } Else {
            $ExceptionMessage = "The workbook $($Path) does not contain any valid worksheets."
        }

        If (!($ExceptionMessage)) {
            $Workbook.Sheets.Item($SelectedSheet).Activate()
            $Script:Cols = $Workbook.ActiveSheet.usedRange.Columns.Count
            $Script:Rows = $Workbook.ActiveSheet.usedRange.Rows.Count
            $Script:Data = $Workbook.ActiveSheet.usedRange.Value2
        }

        $Workbook.Close()
        $ExcelApplication.Quit()
        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheets)) { }
        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook)) { }
        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApplication)) { }
        [GC]::Collect()

        If ($ExceptionMessage) {
            $ExceptionMessage = New-Object System.FormatException $ExceptionMessage
            Throw $ExceptionMessage
        }

    }

    If (Test-Path $Path) {
        ReadData $Path
        $Script:Headers = @()
        $Row = 2
        $HeaderOffset = 1
        If ($HasTitle) {
            $Row = $Row + 2
            $HeaderOffset = 3
        }
        If (!($IsTransposed)) {
            1..$Script:Cols | % { $Script:Headers += $Script:Data[$HeaderOffset,$_] }
            $Row..$Script:Rows | % {
                $CurrentRow = $_
                $Props = $Null
                1..$Script:Cols | % {
                    $Props += [Ordered]@{
                        $($Script:Headers[$_ - 1]) = "$($Script:Data[$CurrentRow,$_])"
                    }
                }
                New-Object PSObject -Property $Props
            }
        } Else {
            If (!($HasTransposeColumnProperty)) { $Row-- }
            $Row..$Script:Rows | % { $Script:Headers += $Script:Data[$_,1] }
            2..$Script:Cols | % {
                $CurrentCol = $_
                $Props = $Null
                $Row..$Script:Rows | % {
                    $Props += [Ordered]@{
                        $($Script:Headers[$_ - $Row]) = "$($Script:Data[$_,$CurrentCol])"
                    }
                }
                New-Object PSObject -Property $Props
            }
        }
    } Else {
        $ExceptionMessage = New-Object System.FormatException "The workbook $($Path) could not be found."
        Throw $ExceptionMessage
    }
}
