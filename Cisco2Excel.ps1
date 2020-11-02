<#
.SYNOPSIS
Cisco2Excel parses the configuration from a Cisco (IOS) device into a Excel file.
.DESCRIPTION
The Cisco2Excel reads a Cisco (IOS) config file and pulls out the configuration into excel.
.PARAMETER CiscoConfig
[REQUIRED] This is the path to the Cisco config/credential file
.PARAMETER SkipFilter 
[OPTIONAL] Set this value to $TRUE for not using Excel Filters.
.\Cisco2Excel.ps1 -CiscoConfig "c:\temp\config.conf"
    Parses a Cisco config file and places the Excel file in the same folder where the config was found.
.\Cisco2Excel.ps1 -CiscoConfig "c:\temp\config.conf" -SkipFilter:$true
    Parses a Cisco config file and places the Excel file in the same folder where the config was found.
    No filters will be auto applied.
.NOTES
Author: Xander Angenent (@XaAng70)
Last Modified: 2020/10/26
#Uses Estimated completion time from http://mylifeismymessage.net/1672/
#Uses Posh-SSH https://github.com/darkoperator/Posh-SSH if reading directly from the firewall
#Uses Function that converts any Excel column number to A1 format https://gallery.technet.microsoft.com/office/Powershell-function-that-88f9f690
#>Param
(
    [Parameter(Mandatory = $true)]
    $CiscoConfig,
    [switch]$SkipFilter = $false
)

Function InitInterface {
    $InitRule = New-Object System.Object;
    $InitRule | Add-Member -type NoteProperty -name Interface -Value ""
    $InitRule | Add-Member -type NoteProperty -name Description -Value ""
    $InitRule | Add-Member -type NoteProperty -name IPadress -Value ""
    $InitRule | Add-Member -type NoteProperty -name IPhelper -Value ""
    $InitRule | Add-Member -type NoteProperty -name "switchport-mode" -Value ""
    $InitRule | Add-Member -type NoteProperty -name "switchport-mode-access-vlan" -Value ""
    $InitRule | Add-Member -type NoteProperty -name "switchport-mode-trunk-native-vlan" -Value ""
    $InitRule | Add-Member -type NoteProperty -name "Switchport-voice-vlan" -Value ""
    $InitRule | Add-Member -type NoteProperty -name "priority-queue" -Value ""
    $InitRule | Add-Member -type NoteProperty -name "spanning-tree" -Value ""
    $InitRule | Add-Member -type NoteProperty -name speed -Value "Auto"
    $InitRule | Add-Member -type NoteProperty -name duplex -Value "Auto"    
    $InitRule | Add-Member -type NoteProperty -name "channel-group" -Value ""  
    $InitRule | Add-Member -type NoteProperty -name "channel-group-mode" -Value ""
    $InitRule | Add-Member -type NoteProperty -name qos -Value ""  
    return $InitRule
}
Function ChangeFontExcelCell ($ChangeFontExcelCellSheet, $ChangeFontExcelCellRow, $ChangeFontExcelCellColumn) {
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).HorizontalAlignment = -4108
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.Size = 18
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.Bold=$True
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.Name = "Cambria"
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.ThemeFont = 1
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.ThemeColor = 4
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.ColorIndex = 55
    $ChangeFontExcelCellSheet.Cells.Item($ChangeFontExcelCellRow, $ChangeFontExcelCellColumn).Font.Color = 8210719
}
Function CreateExcelSheet ($SheetName, $SheetArray) {
    if ($SheetArray) {
        $row = 1
        $Sheet = $workbook.Worksheets.Add()
        $Sheet.Name = $SheetName
        $Column=1
        $excel.cells.item($row,$Column) = $SheetName 
        ChangeFontExcelCell $Sheet $row $Column  
        $row++
        $NoteProperties = SkipEmptyNoteProperties $SheetArray
        foreach ($Noteproperty in $NoteProperties) {
            $excel.cells.item($row,$Column) = $Noteproperty.Name
            $Column++
        }
        $StartRow = $Row
        $row++
        foreach ($rule in $SheetArray) {
            $Column=1
            foreach ($Noteproperty in $NoteProperties) {
                $PropertyString = [string]$NoteProperty.Name
                $Value = $Rule.$PropertyString
                $excel.cells.item($row,$Column) = $Value
                $Column++
            }    
            $row++
        }    
        #No need to filer if there is only one row.
        if (!($SkipFilter) -and ($SheetArray.Count -gt 1)) {
            $RowCount =  $Sheet.UsedRange.Rows.Count
            $ColumCount =  $Sheet.UsedRange.Columns.Count
            $ColumExcel = Convert-NumberToA1 $ColumCount
            $Sheet.Range("A$($StartRow):$($ColumExcel)$($RowCount)").AutoFilter() | Out-Null
        }
        #Use autoFit to expand the colums
        $UsedRange = $Sheet.usedRange                  
        $UsedRange.EntireColumn.AutoFit() | Out-Null
    }
}
#Function from https://gallery.technet.microsoft.com/office/Powershell-function-that-88f9f690
Function Convert-NumberToA1 { 
    <# 
    .SYNOPSIS 
    This converts any integer into A1 format. 
    .DESCRIPTION 
    See synopsis. 
    .PARAMETER number 
    Any number between 1 and 2147483647 
    #> 
     
    Param([parameter(Mandatory=$true)] 
          [int]$number) 
   
    $a1Value = $null 
    While ($number -gt 0) { 
      $multiplier = [int][system.math]::Floor(($number / 26)) 
      $charNumber = $number - ($multiplier * 26) 
      If ($charNumber -eq 0) { $multiplier-- ; $charNumber = 26 } 
      $a1Value = [char]($charNumber + 64) + $a1Value 
      $number = $multiplier 
    } 
    Return $a1Value 
  }
Function GetSubnetCIDR ([string]$Subnet,[IPAddress]$SubnetMask) {
    $binaryOctets = $SubnetMask.GetAddressBytes() | ForEach-Object { [Convert]::ToString($_, 2) }
    $SubnetCIDR = $Subnet + "/" + ($binaryOctets -join '').Trim('0').Length
    return $SubnetCIDR
}
#Function SkipEmptyNoteProperties ($SkipEmptyNotePropertiesArray)
#This function Loopt through all available noteproperties and checks if it is used.
#If it is not used the property will not be returned as it is not needed in the export.
Function SkipEmptyNoteProperties ($SkipEmptyNotePropertiesArray) {
    $ReturnNoteProperties = [System.Collections.ArrayList]@()
    $SkipNotePropertiesOrg = $SkipEmptyNotePropertiesArray | get-member -Type NoteProperty
    foreach ($SkipNotePropertieOrg in $SkipNotePropertiesOrg) {
        foreach ($SkipEmptyNotePropertiesMember in $SkipEmptyNotePropertiesArray) {
            $NotePropertyFound = $False
            $SkipNotePropertiePropertyString = [string]$SkipNotePropertieOrg.Name
            if ($SkipEmptyNotePropertiesMember.$SkipNotePropertiePropertyString) { 
                $NotePropertyFound = $True
                break;
            }
        }
        If ($NotePropertyFound) { $ReturnNoteProperties.Add($SkipNotePropertieOrg) | Out-Null  }
    }

    return $ReturnNoteProperties
}

$startTime = get-date 
$date = Get-Date -Format yyyyMMddHHmm
Clear-Host
Write-Output "Started script"
#Clear 5 additional lines for the progress bar
$I=0
DO {
    Write-output ""
    $I++
} While ($i -le 4)
If ($SkipFilter) {
    Write-Output "SkipFilter parmeter is set to True. Skipping filter function in Excel."
}
if (!(Test-Path $CiscoConfig)) {
    Write-Output "File $CiscoConfig not found. Aborting script."
    exit 1
}
$loadedConfig = Get-Content $CiscoConfig
$Counter=0
$workingFolder = Split-Path $CiscoConfig;
$fileName = Split-Path $CiscoConfig -Leaf;
$fileName = (Get-Item $CiscoConfig).Basename
$ExcelFullFilePad = "$workingFolder\$fileName"
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$FirstSheet = $workbook.Worksheets.Item(1) 
$FirstSheet.Name = $FileName
$FirstSheet.Cells.Item(1,1)= 'Cisco Configuration'
$MergeCells = $FirstSheet.Range("A1:G1")
$MergeCells.Select() | Out-Null
$MergeCells.MergeCells = $true
$FirstSheet.Cells(1, 1).HorizontalAlignment = -4108
$FirstSheet.Cells.Item(1,1).Font.Size = 18
$FirstSheet.Cells.Item(1,1).Font.Bold=$True
$FirstSheet.Cells.Item(1,1).Font.Name = "Cambria"
$FirstSheet.Cells.Item(1,1).Font.ThemeFont = 1
$FirstSheet.Cells.Item(1,1).Font.ThemeColor = 4
$FirstSheet.Cells.Item(1,1).Font.ColorIndex = 55
$FirstSheet.Cells.Item(1,1).Font.Color = 8210719
$InterfaceSwitch=$False
$MaxCounter=$loadedConfig.count
$InterfaceList = [System.Collections.ArrayList]@()
$RouterTable = [System.Collections.ArrayList]@()
foreach ($Line in $loadedConfig) {
    $Proc = $Counter/$MaxCounter*100
    $elapsedTime = $(get-date) - $startTime 
    if ($Counter -eq 0) { $estimatedTotalSeconds = $MaxCounter/ 1 * $elapsedTime.TotalSecond }
    else { $estimatedTotalSeconds = $MaxCounter/ $counter * $elapsedTime.TotalSeconds }
    $estimatedTotalSecondsTS = New-TimeSpan -seconds $estimatedTotalSeconds
    $estimatedCompletionTime = $startTime + $estimatedTotalSecondsTS    
    Write-Progress -Activity "Parsing config file. Estimate completion time $estimatedCompletionTime" -PercentComplete ($Proc)
    $Counter++
    $Configline=$Line.Trim() -replace '\s+',' '
    $ConfigLineArray = $Configline.Split(" ")    
    switch($ConfigLineArray[0]) {
        "" {
            #Do nothing
        }
        "Channel-group" {
            $Interface | Add-Member -MemberType NoteProperty -Name channel-group -Value $ConfigLineArray[1] -force
            $Interface | Add-Member -MemberType NoteProperty -Name channel-group-mode -Value $ConfigLineArray[3] -force              
        }
        "description" {
            if ($ConfigLineArray.Count -eq 2) { $Value = $ConfigLineArray[1]}
            else {
                $Value = $ConfigLineArray[1]
                For ($ConfigLineArrayCount=2; $ConfigLineArrayCount -le $ConfigLineArray.Count; $ConfigLineArrayCount++) {
                $Value = $Value + " " + $ConfigLineArray[$ConfigLineArrayCount]
                }
            }           
            $Interface | Add-Member -MemberType NoteProperty -Name $ConfigLineArray[0] -Value $Value -force
        }
        "hostname" {
            $Hostname = $ConfigLineArray[1]
        }
        "interface" {
            $Interface = InitInterface
            $Interface | Add-Member -MemberType NoteProperty -Name "Interface" -Value $ConfigLineArray[1] -force
            $InterfaceSwitch=$true
        }
        "ip" {
            if ($InterfaceSwitch) {
                switch($ConfigLineArray[1]) {
                    "address" {
                        $Value = GetSubnetCIDR $ConfigLineArray[2] $ConfigLineArray[3] 
                        $Interface | Add-Member -MemberType NoteProperty -Name IPadress -Value $Value -force
                    }
                    "helper-address" {
                        $Value = $Interface.IPhelper
                        if ($Value -eq "") {
                            $Value = $ConfigLineArray[2]
                        }
                        else {
                            $Value = $Value + "," + $ConfigLineArray[2]
                        }
                        $Interface | Add-Member -MemberType NoteProperty -Name IPhelper -Value $Value -force
                    }

                }
            }
            else {
                if ($ConfigLineArray[1] -eq "route") {
                    $Value = GetSubnetCIDR $ConfigLineArray[2] $ConfigLineArray[3]
                    $Route = New-Object System.Object;
                    $Route | Add-Member -type NoteProperty -name Network -Value $Value
                    $Route | Add-Member -type NoteProperty -name Gateway -Value $ConfigLineArray[4]
                    $RouterTable.Add($Route) | Out-Null
                }
            }
        }
        "no" {
            if ($InterfaceSwitch) {
                if ($ConfigLineArray[1] -eq "ip") { $Interface | Add-Member -MemberType NoteProperty -Name IPadress -Value "no ip address" -force }
            }
        }
        "spanning-tree" {
            if ($InterfaceSwitch) {
                $Interface | Add-Member -MemberType NoteProperty -Name spanning-tree -Value $ConfigLineArray[1] -force 
            }
        }
        "switchport" {
            switch ($ConfigLineArray[1]){
                "Access" {
                    $Interface | Add-Member -MemberType NoteProperty -Name switchport-mode-access-vlan -Value $ConfigLineArray[3] -force 
                }
                "mode" {
                    $Interface | Add-Member -MemberType NoteProperty -Name switchport-mode -Value $ConfigLineArray[2] -force 
                }
                "trunk" {
                    $Interface | Add-Member -MemberType NoteProperty -Name switchport-mode-trunk-native-vlan -Value $ConfigLineArray[4] -force                    
                }
                "Voice" {
                    $Interface | Add-Member -MemberType NoteProperty -Name Switchport-voice-vlan -Value $ConfigLineArray[3] -force 
                }
            }
        }
        "!" {
            if ($InterfaceSwitch) {
                $InterfaceList.Add($Interface) | Out-Null
                $InterfaceSwitch=$False
            }
        }
        default {
            if ($InterfaceSwitch) {
                $Interface | Add-Member -MemberType NoteProperty -Name $ConfigLineArray[1] -Value $ConfigLineArray[2] -force
            }
        }
    }
}
#make sure that the first sheet that is opened by Excel is the global sheet.
CreateExcelSheet "Interfaces" $Interfacelist 
CreateExcelSheet "RoutingTable" $RouterTable
$FirstSheet.Activate()
$FirstSheet.Cells.Item(2,1) = 'Excel Creation Date'
$FirstSheet.Cells.Item(2,2) = $Date
$FirstSheet.Cells.Item(2,2).numberformat = "00"
$FirstSheet.Cells.Item(3,1) = 'Config Creation Date'
$LastConfigLine = $loadedConfig[2]
$FirstSheet.Cells.Item(3,2) = $LastConfigLine 
#$FirstSheet.Cells.Item(3,2).numberformat = "00" 
$FirstSheet.Cells.Item(4,1) = 'Last NVRAM write Date'
$LastNVRAMConfigLine = $loadedConfig[3]
$FirstSheet.Cells.Item(4,2) = $LastNVRAMConfigLine
#$FirstSheet.Cells.Item(4,2).numberformat = "00"
$FirstSheet.Cells.Item(5,1) = 'Hostname'
$FirstSheet.Cells.Item(5,2) = $Hostname
$UsedRange = $FirstSheet.usedRange                  
$UsedRange.EntireColumn.AutoFit() | Out-Null
Write-Output "Writing Excelfile $ExcelFullFilePad.xls"
$workbook.SaveAs($ExcelFullFilePad)
$excel.Quit()
$elapsedTime = $(get-date) - $startTime
$Minutes = $elapsedTime.Minutes
$Seconds = $elapsedTime.Seconds
Write-Output "Script done in $Minutes Minute(s) and $Seconds Second(s)."