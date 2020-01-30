Param
(
    [Parameter(Mandatory = $true)]
    $CiscoConfig
)

Function InitInterface {
    $InitRule = New-Object System.Object;
    $InitRule | Add-Member -type NoteProperty -name Interface -Value ""
    $InitRule | Add-Member -type NoteProperty -name Description -Value ""
    $InitRule | Add-Member -type NoteProperty -name IPadress -Value ""
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
Function GetSubnetCIDR ([string]$Subnet,[IPAddress]$SubnetMask) {
    $binaryOctets = $SubnetMask.GetAddressBytes() | ForEach-Object { [Convert]::ToString($_, 2) }
    $SubnetCIDR = $Subnet + "/" + ($binaryOctets -join '').Trim('0').Length
    return $SubnetCIDR
}
Function CreateExcelSheet ($SheetName, $SheetArray) {
    if ($SheetArray) {
        $row = 2
        $Sheet = $workbook.Worksheets.Add()
        $Sheet.Name = $SheetName
        $Column=1
        $NoteProperties = $SheetArray | get-member -Type NoteProperty
        foreach ($Noteproperty in $NoteProperties) {
            $excel.cells.item(1,$Column) = $Noteproperty.Name
            $Column++
        }
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
        #Use autoFit to expand the colums
        $UsedRange = $Sheet.usedRange                  
        $UsedRange.EntireColumn.AutoFit() | Out-Null
    }
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
$MergeCells.Select()
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
$InterfaceList = @()
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
            #Do noting
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
                $Value = GetSubnetCIDR $ConfigLineArray[2] $ConfigLineArray[3] 
                $Interface | Add-Member -MemberType NoteProperty -Name IPadress -Value $Value -force
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
                $InterfaceList += $Interface
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