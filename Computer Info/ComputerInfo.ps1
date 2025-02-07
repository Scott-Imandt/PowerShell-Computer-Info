Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$currentDirectory = Get-Location
Write-Host $currentDirectory
$xamlPath = Join-Path $currentDirectory "\MainWindow.xaml"
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


$Columns=@(
    'Name'
    'DriverName'
    'PortName'
)



$var_DataGrid_Printers.FontSize= '18'
$DeviceDataTable=New-Object System.Data.DataTable
[void]$DeviceDataTable.Columns.AddRange($Columns)


$var_Button_Submit.Add_Click({
    
    $ComputerName = $var_TextBox_ComputerName.Text

    try{
        $DeviceDataTable.Clear()
        $var_DataGrid_Printers.Clear()
        $Devices = get-wmiobject -class win32_printer -computername $ComputerName -property "Name, DriverName, PortName" | Select-Object $Columns
    
        if ($Devices) {

            
            foreach($Device in $Devices){
                $Entry=@()
                foreach($Column in $Columns){
                    $Entry+=$Device.$Column
                }
                [void]$DeviceDataTable.Rows.Add($Entry)
            }
         
            $var_DataGrid_Printers.ItemsSource=$DeviceDataTable.DefaultView
            $var_DataGrid_Printers.IsReadOnly=$true
            $var_DataGrid_Printers.GridLinesVisibility="Horizontal"
        
        } else {
            [System.Windows.Forms.MessageBox]::Show("No Devices found on $ComputerName.")      
        }
    
    }catch{
    
        [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
    
    }

})




$form1.showDialog()




