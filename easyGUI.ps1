### EasyGUI to quickly generate an Excel file, or simply to view or print data ###


### Modifiable Variables ###

$Icon = $PSScriptRoot + "\Examplelogo.bmp"     ## path to Icon for GUI, you do not need to include one
$csvpath = $PSScriptRoot + ".\Example.csv"    ## Path to your CSV file 


# Load Assemblies

[System.Reflection.Assembly]::LoadwithPartialName("System.Windows.Forms") | Out-Null
[System.Reflection.Assembly]::LoadwithPartialName("System.Drawing") | Out-Null


Function Filter-Type {
    $global:Filter = "Type"

    $SecondCheckBox.Items.Clear()
    $ThirdCheckBox.Items.Clear()
    if ($SecondCheckBox.Items.Count -eq 0) {

        $SecondCheckBoxItems = $Csv | ForEach-Object {$_.Type} | Select-Object -Unique
        for ($i = 0; $i -lt $SecondCheckBoxItems.count; $i++) {
            $SecondCheckBox.Items.Add($SecondCheckBoxItems[$i])
            
        }
        $Type = $SecondCheckBoxItems
        $Box = $SecondCheckBox          
    }

}


Function Filter-Building {
    $global:filter = "Building"    
    $SecondCheckBox.Items.Clear() 
    $ThirdCheckBox.Items.Clear()     
    if ($SecondCheckBox.Items.Count -eq 0) {

        $SecondCheckBoxItems = $Csv | ForEach-Object {$_.Building} | Select-Object -Unique
        for ($i = 0; $i -lt $SecondCheckBoxItems.count; $i++) {
            $SecondCheckBox.Items.Add($SecondCheckBoxItems[$i])
        
        }        
    }    
}



Function Filter-Area {
    $Global:filter = "Area"
    $SecondCheckBox.Items.Clear()
    $ThirdCheckBox.Items.Clear()     
    if ($SecondCheckBox.Items.Count -eq 0) {

        $SecondCheckBoxItems = $Csv | ForEach-Object {$_.Area} | Select-Object -Unique
        for ($i = 0; $i -lt $SecondCheckBoxItems.count; $i++) {
            $SecondCheckBox.Items.Add($SecondCheckBoxItems[$i])
        
        }        
    }    
}


Function Filter-Custom {
    $global:filter = "SubType"
    $SecondCheckBox.Items.Clear()
    $ThirdCheckBox.Items.Clear()       
    if ($SecondCheckBox.Items.Count -eq 0) {

        $SecondCheckBoxItems = $Csv | ForEach-Object {$_.SubType} | Select-Object -Unique
        for ($i = 0; $i -lt $SecondCheckBoxItems.count; $i++) {
            $SecondCheckBox.Items.Add($SecondCheckBoxItems[$i])
        
        }        
    }
    
    $global:filter2 = "Building"  
    if ($ThirdCheckBox.Items.Count -eq 0) {

        $ThirdCheckBoxItems = $Csv | ForEach-Object {$_.Building} | Select-Object -Unique
        for ($i = 0; $i -lt $ThirdCheckBoxItems.count; $i++) {
            $ThirdCheckBox.Items.Add($ThirdCheckBoxItems[$i])
        
        }        
    }    

        
}

Function Grid-View {


    Foreach ($i in $SecondCheckBox.CheckedItems) {
    
        $x += $i +"|"
                    
    }
    $y = $x.Substring(0, $x.length - 1) 

    Foreach ($i in $ThirdCheckBox.CheckedItems) {
    
        $n += $i +"|"
                    
    }

    If ($n -ne $null) {
        $z = $n.Substring(0, $n.length - 1) 
    }
    $csv | Where-Object {$_.$filter -match $y -and $_.$filter2 -match $z} | Out-GridView -erroraction 'Stop'

}

Function Save-AsExcel {

    Foreach ($i in $SecondCheckBox.CheckedItems) {
    
        $x += $i +"|"
                    
    }
    $y = $x.Substring(0, $x.length - 1) 

    Foreach ($i in $ThirdCheckBox.CheckedItems) {
    
        $n += $i +"|"
                    
    }
    If ($n -ne $null) {
        $z = $n.Substring(0, $n.length - 1) 
    }

    $csv1 = $csv | Where-Object {$_.$filter -Match $y -and $_.$filter2 -match $z}

    $excel = New-Object -ComObject excel.application

    $workbook = $excel.Workbooks.Add()

    $excel.DisplayAlerts = $False

    $devicesheet = $workbook.Worksheets.Item(1)
    $devicesheet.Name = "1"
    $devicesheet.Cells.Item(1,1) = "Device Information"
    $devicesheet.Cells.Item(1,1).Font.Color = 8210719

    $headerrow = 2
    $column = 1

    $csv1 | Get-Member -Type NoteProperty | ForEach-Object {        
        $devicesheet.Cells.Item($headerrow, $column) = $_.Name
        $column++
    }

    $row = 3

    foreach($line in $csv1)
    {        
        $properties = $line | Get-Member -MemberType Properties

        for($i=0; $i -lt $properties.Count;$i++)
        {
            $column = $properties[$i]
            $columnvalue = $line | Select -ExpandProperty $column.Name
            $devicesheet.Cells.Item($row, $i+1) = $columnvalue 
        }

        $row++
    } 

    $usedRange = $devicesheet.UsedRange
    $usedRange.EntireColumn.AutoFit() | Out-Null

    $sfg = New-Object system.windows.forms.savefiledialog
    $sfg.Title = 'Select a file'
    $sfg.InitialDirectory = $env:USERPROFILE + "\Desktop"
    $sfg.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
    $sfg.ShowHelp = $true
    $sfg.showdialog()
    $sfg.filename

    $workbook.SaveAs($sfg.FileName)

    $excel.Quit()
}


Function Send-ToPrinter {

    Foreach ($i in $SecondCheckBox.CheckedItems) {
    
        $x += $i +"|"
                    
    }
    $y = $x.Substring(0, $x.length - 1) 

    Foreach ($i in $ThirdCheckBox.CheckedItems) {
    
        $n += $i +"|"
                    
    }

    If ($n -ne $null) {
        $z = $n.Substring(0, $n.length - 1) 
    }

    $csv | Where-Object {$_.$filter -match $y -and $_.$filter2 -match $z} | Format-Table | Out-Printer


}




###


$Form = New-Object System.Windows.Forms.Form
$Form.width = 800
$Form.height = 500
$Form.Font = New-Object System.Drawing.Font("Helvetica", 13)
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Fixed3D
$Form.BackColor = "#0d0c0c"
$Form.ForeColor = "LawnGreen"
$Form.Text = "Easy GUI for CSV"
$Form.MaximumSize = New-Object System.Drawing.Size(1920,1080)
$Form.startposition = "centerscreen"


If ((Test-Path $Icon -PathType Leaf) -eq $True) {

    $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($Icon)

}

### Add buttons ###

# First Selection #

$FirstButton1 = New-Object System.Windows.Forms.RadioButton
$FirstButton1.Text = "Type Only"
$FirstButton1.Location = New-Object System.Drawing.Size(30,30)
$FirstButton1.Size = New-Object System.Drawing.Size(130,50)
$FirstButton1.Add_Click({Filter-Type})

$FirstButton2 = New-Object System.Windows.Forms.RadioButton
$FirstButton2.Text = "Building Only"
$FirstButton2.Location = New-Object System.Drawing.Size(30,80)
$FirstButton2.Size = New-Object System.Drawing.Size(130,50)
$FirstButton2.Add_Click({Filter-Building})

$FirstButton3 = New-Object System.Windows.Forms.RadioButton
$FirstButton3.Text = "Area Only"
$FirstButton3.Location = New-Object System.Drawing.Size(30,130)
$FirstButton3.Size = New-Object System.Drawing.Size(130,50)
$FirstButton3.Add_Click({Filter-Area})

$FirstButton4 = New-Object System.Windows.Forms.RadioButton
$FirstButton4.Text = "Custom"
$FirstButton4.Location = New-Object System.Drawing.Size(30,180)
$FirstButton4.Size = New-Object System.Drawing.Size(130,50)
$FirstButton4.Add_Click({Filter-Custom})

$FirstGroupBox = New-Object System.Windows.Forms.GroupBox
$FirstGroupBox.Text = "Filter By"
$FirstGroupBox.Location = New-Object System.Drawing.Size(50,80) 
$FirstGroupBox.Size = New-Object System.Drawing.Size(170,270)
$FirstGroupBox.ForeColor = "LawnGreen"

$SecondCheckBox = New-Object System.Windows.Forms.CheckedListBox
$SecondCheckBox.BackColor = "#0d0c0c"
$SecondCheckBox.Forecolor = "LawnGreen"
$SecondCheckBox.Location = New-Object System.Drawing.Size(240,100)
$SecondCheckBox.CheckOnClick = $True
$SecondCheckBox.Width = 160
$SecondCheckBox.Height = 190
$SecondCheckBox.Sorted = $True

$ThirdCheckBox = New-Object System.Windows.Forms.CheckedListBox
$ThirdCheckBox.BackColor = "#0d0c0c"
$ThirdCheckBox.Forecolor = "LawnGreen"
$ThirdCheckBox.Location = New-Object System.Drawing.Size(400,100)
$ThirdCheckBox.CheckOnClick = $True
$ThirdCheckBox.Width = 120
$ThirdCheckBox.Height = 190
$ThirdCheckBox.Sorted = $True

$OutViewGridButton = New-Object System.Windows.Forms.Button
$OutViewGridButton.Location = New-Object System.Drawing.Size(600,100)
$OutViewGridButton.Text = "Grid Pop-Up"
$OutViewGridButton.Size = New-Object System.Drawing.Size(150,50)
$OutViewGridButton.Add_Click({Grid-View [ref]$SecondCheckBox,[ref]$ThirdCheckBox})


$SaveAsExcelButton = New-Object System.Windows.Forms.Button
$SaveAsExcelButton.Location = New-Object System.Drawing.Size(600,155)
$SaveAsExcelButton.Text = "Save As Excel"
$SaveAsExcelButton.Size = New-Object System.Drawing.Size(150,50)
$SaveAsExcelButton.Add_Click({Save-AsExcel [ref]$SecondCheckBox, [ref]$ThirdCheckBox})


#Default printer
$Printer = Get-WmiObject -Class win32_printer | Where-Object {$_.Default -eq "True"}

$SendPrinter = New-Object System.Windows.Forms.Button
$SendPrinter.Location = New-Object System.Drawing.Size(600,210)
$SendPrinter.Text = "Print to " + $Printer.Name
$SendPrinter.Size = New-Object System.Drawing.Size(150,50)
$SendPrinter.Add_Click({Send-ToPrinter [ref]$SecondCheckBox, [ref]$ThirdCheckBox})


Function easyGUI {

    $FirstGroupBox.Controls.Add($FirstButton1)
    $FirstGroupBox.Controls.Add($FirstButton2)
    $FirstGroupBox.Controls.Add($FirstButton3)
    $FirstGroupBox.Controls.Add($FirstButton4)

    $Form.Controls.Add($OutViewGridButton)
    $Form.Controls.Add($SaveAsExcelButton)
    $Form.Controls.Add($SendPrinter)

    $Form.Controls.Add($FirstGroupBox)
    $Form.Controls.Add($SecondCheckBox)
    $Form.Controls.Add($ThirdCheckBox)

    If ((Test-Path $Icon) -eq $True)
    {
        $Form.Icon = [System.drawing.icon]::ExtractAssociatedIcon($Icon)
    }

    $Form.ShowDialog()
}


$Csv = Import-Csv -Path $Csvpath


easyGUI