Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$currentDirectory = Get-Location
Write-Host $currentDirectory
$xamlPath = Join-Path $currentDirectory "\MainWindow.xaml"
#$xamlPath = "D:\Computer Info\MainWindow.xaml"
$inputXAML=Get-Content -Path $xamlPath -Raw
$inputXAML=$inputXAML -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
[XML]$XAML=$inputXAML

$reader = New-Object System.Xml.XmlNodeReader $XAML
try{
    $form1=[Windows.Markup.XamlReader]::Load($reader)
}catch{
    Write-Host $_.Exception
    throw
}

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    try{
        Set-Variable -Name "var_$($_.Name)" -Value $form1.FindName($_.Name) -ErrorAction Stop
    }catch{
        throw
    }
}

$form1.Icon = Join-Path $currentDirectory "\Computer info Icon.ico"


$ColumnsPrinters=@(
    'Name'
    'DriverName'
    'PortName'
)

$ColumnsUSB=@(
    'Manufacturer'
    'Description'
    'Service'
    'PNPClass'
    'DeviceID'

)

$var_DataGrid_Printers.FontSize= '18'
$PrinterDataTable=New-Object System.Data.DataTable
[void]$PrinterDataTable.Columns.AddRange($ColumnsPrinters)

$var_DataGrid_USB.FontSize= '18'
$USBDataTable=New-Object System.Data.DataTable
[void]$USBDataTable.Columns.AddRange($ColumnsUSB)

$Global:ComputerName = $null


$var_Button_Submit.Add_Click({
    
    $Global:ComputerName = $var_TextBox_ComputerName.Text

    try{
        $PrinterDataTable.Clear()
        $var_DataGrid_Printers.Clear()
        $USBDataTable.Clear()
        $var_DataGrid_USB.Clear()
        $var_Button_DIR.Visibility = 'Hidden'
        
        $Printers = get-wmiobject -class win32_printer -computername $ComputerName -property "Name, DriverName, PortName" | Select-Object $ColumnsPrinters
        
        $USBs = gwmi Win32_USBControllerDevice -computername $ComputerName |%{[wmi]($_.Dependent)} | Sort Manufacturer,Description,PNPClass,DeviceID | Select-Object $ColumnsUSB

        if($Printers -or $USBs){
            $var_Button_DIR.Visibility = 'Visible'
        }
        else{
            
            [System.Windows.Forms.MessageBox]::Show("No USB or Printer Devices found on $ComputerName.")
            return
        }

        if ($Printers) {

            foreach($Printer in $Printers){
                $Entry=@()
                foreach($Column in $ColumnsPrinters){
                    $Entry+=$Printer.$Column
                }
                [void]$PrinterDataTable.Rows.Add($Entry)
            }
         
            $var_DataGrid_Printers.ItemsSource=$PrinterDataTable.DefaultView
            $var_DataGrid_Printers.IsReadOnly=$true
            $var_DataGrid_Printers.GridLinesVisibility="Horizontal"
        
        } else {
            [System.Windows.Forms.MessageBox]::Show("No Printers found on $ComputerName.")      
        }

        if ($USBs) {
 
            foreach($USB in $USBs){
                $Entry=@()
                foreach($Column in $ColumnsUSB){
                    $Entry+=$USB.$Column
                }
                [void]$USBDataTable.Rows.Add($Entry)
            }
         
            $var_DataGrid_USB.ItemsSource=$USBDataTable.DefaultView
            $var_DataGrid_USB.IsReadOnly=$true
            $var_DataGrid_USB.GridLinesVisibility="Horizontal"
        
        } else {
            [System.Windows.Forms.MessageBox]::Show("No USB Devices found on $ComputerName.")      
        }
    
    }catch{
    
        [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
    
    }


})

$var_Button_DIR.Add_Click({
    
    try{
        Invoke-Item "\\$ComputerName\c$" -ErrorAction Stop
    }
    catch{
        [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
        
    }

})


$form1.showDialog()




