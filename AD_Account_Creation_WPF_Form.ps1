##########
# Edit these variables, based on what you have in your requirments 

#What OU would you like these accounts created in?
$path = ""

#how long would you like the mnemonic to be?
$length = 6

##########

#-------------------------------------------------------------#
# This script will create the AD accounts when uploading with an CSV
# By: Fabio De Oliveira
# Gui created with PoshGui.com
#-------------------------------------------------------------#


#### Bugs ####

# - Error coming from Clean-Up-Blank-Rows-DataGrid after clearing blank rows  (Collection was modified; enumeration operation might not execute.) 
    # Ignore this if you see this error message, it has to do with the script trying to clear blank lines


#----Initial Declarations-------------------------------------#
#-------------------------------------------------------------#
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
Add-Type -AssemblyName PresentationCore, PresentationFramework


#XML Window, this is where you modify the GUI (Add Buttons, Labels, Etc)
$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="950" Height="650">
<Grid>
<DataGrid HorizontalAlignment="Left" VerticalAlignment="Top" Width="700" Height="400" Margin="100,100,0,0" Name="DataGrid1"/>
<Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="Active Directory Account Creation" Margin="100,10,0,0" Name="Label1"/>
<Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="NOTE: Every CSV requires the following Headers or you will not be able to open it" Margin="180,40,0,0" Name="Label2"/>
<Label HorizontalAlignment="Left" VerticalAlignment="Top" Content="FirstName      LastName     Title" Margin="180,60,0,0" Name="Label3"/>
<Button Content="Load CSV" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="100,40,0,0" Name="LoadCSVButton"/>
<Button Content="Create AD Accounts" HorizontalAlignment="Left" VerticalAlignment="Top" Width="150"  Margin="100,520,0,0" Name="CreateADButton"/>
<Button Content="Update Duplicate's Password" HorizontalAlignment="Left" VerticalAlignment="Top" Width="150"  Margin="500,520,0,0" Name="UpdateDupPassword"/>
<Button Content="Export CSV" HorizontalAlignment="Left" VerticalAlignment="Top" Width="150"  Margin="100,560,0,0" Name="ExportCSVButton"/>
</Grid></Window>
"@

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)
[xml]$xml = $Xaml
$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }

#-------------------------------------------------------------#
#----Functions------------------------------------------------#
#-------------------------------------------------------------#

#Function to open File Dialog (This will open a file explorer window which will allow you to load in a file)
function RETURN-IMPORT-CSV_OpenFileDialog {
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'SpreadSheet (*.csv)|*.csv'
    }
    if($FileBrowser.ShowDialog() -ne "OK") {
      return $null
    }else{
        $csv_file = import-csv ($FileBrowser.FileName)
        return $csv_file
    }
}

#Function which will confirm that the following CSV loaded, has the proper headers. (Return True if Match))
function RETURN-BOOL_Confirm-Headers-in-CSV ($csv_file){
    $headers_in_csv = ($csv_file | Get-Member -MemberType NoteProperty).Name
    $header_list = @("FirstName","LastName","Title")
    $diff = Compare-Object $headers_in_csv $header_list -ExcludeDifferent -IncludeEqual
    if ($diff.count -eq $header_list.count){
        return $true
    }else{
        return $false
    }
}

#Function which will load the CSV data into the Datagrid 
function CSV-to-DataGrid{
   foreach ($person in $csv){
       $row = $table.NewRow()
       $row.Check = $true
       $row.FirstName = $person.'FirstName'
       $row.Lastname = $person.'LastName'
       $row.Title = $person.'Title'
       $table.Rows.Add($row)
   }
}

#Function which will delete any blank rows
function Clean-Up-Blank-Rows-DataGrid{
    foreach ($row in $table){
        if ([string]::IsNullOrEmpty($row.FirstName) -and [string]::IsNullOrEmpty($row.LastName)){
            $table.Rows.Remove($row)
        }
    }
}

#Function will check the first and last name and will return the AD mnemonic and description if duplicate exists
function Duplicates-Populate-Mnmemonic-DEscription-Field{
    foreach ($row in $table){
        $firstname = $row.FirstName
        $lastname = $row.LastName
        $duplicate_found = [bool](get-aduser -filter {givenname -eq $firstname -and surname -eq $lastname})
        if ($duplicate_found -eq $true){
            $mnemonic = (get-aduser -filter {givenname -eq $firstname -and surname -eq $lastname}) | Select-Object -ExpandProperty SamAccountName
            $description = (Get-AdUser -filter {givenname -eq $firstname -and surname -eq $lastname} -Properties Description | Select-Object -ExpandProperty Description)
            $row.Title = $description
            $row.Mnemonic = $mnemonic
            $row.Duplicate = $true

        }
    }
}

function NonDuplicates-Populate-Mnmemonic-Description-Field{
    foreach ($row in $table){
        $firstname = $row.FirstName
        $lastname = $row.LastName
        $duplicate_found = [bool](get-aduser -filter {givenname -eq $firstname -and surname -eq $lastname})
        if ($duplicate_found -eq $false){
            $row.Duplicate = $false
            $row.Mnemonic = Generate-Mnemonic $firstname $lastname
        }
    }
}

function Generate-Mnemonic ($firstname,$lastname){
    $count = 0
    $firstname = $firstname -replace '[^a-zA-Z]'
    $lastname = $lastname -replace '[^a-zA-Z]'

    DO{ $mnemonic = ($firstname[0] + ($lastname.replace(' ' , '').Substring(0,[System.Math]::Min($length, ($lastname | Measure-Object -Character -IgnoreWhiteSpace | Select -ExpandProperty Characters))) + ("{0:D2}" -f $count | where {$_ -ne "00"}))).Tolower()
        $count++}
    UNTIL (([bool] (Get-ADUser -Filter { SamAccountName -eq $mnemonic })) -eq $False)
    return $mnemonic
}

function Assign-Complex-Password{
    foreach ($row in $table){
        $row.Password = Generate-Password
    }
}

function Generate-Password{
    $number = Get-Random -Minimum 10 -Maximum 99 
    $password = New-Object -TypeName PSObject
    $password | Add-Member -MemberType ScriptProperty -Name "Password" -Value { ("!@#*123456789ABCDEFGHJKLMNPQRSTUVWXYZ_abcdefghijkmnprstuvwxyz".tochararray() | sort {Get-Random})[0..12] -join '' } | Select -ExpandProperty Password
    $password = $password | Select-Object -Property * | Select -ExpandProperty Password
    return $password + $number
}


function Create-AD-Accounts-Button{
    foreach ($row in $table){
        if ($row.Check -eq $true){
            if ($row.Duplicate -eq $false){
                #Include confirmation account was created
                Create-AD-Account ($row.FirstName) ($row.LastName) ($row.Mnemonic) ($row.Description) ($row.Password) ($path)
            }
        }  
    }
}

function Update-Duplicate-Passwords{
    foreach ($row in $table){
        if ($row.Check -eq $true){
            if ($row.Duplicate -eq $true){
                #Include confirmation, password has been updated
                Set-ADAccountPassword -Identity ($row.Mnemonic) -Reset -NewPassword (ConvertTo-SecureString -AsPlainText ($row.Password) -Force)
                Enable-ADAccount -Identity ($row.Mnemonic)
            }
        }  
    }
}

function Create-AD-Account ($firstname, $lastname, $mnemonic, $title, $password, $path){
    New-ADUser -Name "$lastname, $firstname" `
    -displayName "$lastname, $firstname" `
    -givenName $firstname -Surname $lastname `
    -SamAccountName $mnemonic `
    -Description $title `
    -Title $title `
    -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
    -Path $path
    -Enabled $true
}

function Export-Data-As-Csv{
    Add-Type -AssemblyName System.Windows.Forms
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|Excel Worksheet (*.xls)|*.xls|All Files (*.*)|*.*"
        $dlg.SupportMultiDottedExtensions = $true;
        $dlg.InitialDirectory = "C:\temp\"
        $dlg.CheckFileExists = $false;

    if($dlg.ShowDialog() -eq 'Ok'){
        $table | Export-Csv -Path "$($dlg.filename)" -NoTypeInformation
    }
}

#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#
#("FirstName","LastName","EndDate","Mnemonic","Description","Password","Meditech Template","Type")
#Create Data Table
$table = New-Object system.Data.DataTable 'DataTable'
$newcol = New-Object system.Data.DataColumn Check,([bool]); $table.columns.add($newcol)
$newcol = New-Object system.Data.DataColumn Duplicate,([String]); $table.columns.add($newcol);
$newcol = New-Object system.Data.DataColumn FirstName,([string]); $table.columns.add($newcol)
$newcol = New-Object system.Data.DataColumn LastName,([string]); $table.columns.add($newcol)
$newcol = New-Object system.Data.DataColumn Title,([string]); $table.columns.add($newcol)
$newcol = New-Object system.Data.DataColumn Mnemonic,([string]); $table.columns.add($newcol)
$newcol = New-Object system.Data.DataColumn Password,([string]); $table.columns.add($newcol)



#Import Data Table to the On-Screen DataGrid
$DataGrid1.ItemsSource = $table.DefaultView

#Button "Load CSV"
$LoadCSVButton.Add_Click({
    #Open file dialog window to choose CSV file
    $csv = RETURN-IMPORT-CSV_OpenFileDialog

    #if CSV is empty don't run this
    if ($csv -ne $null){
        #Confirms if CSV has the proper headers - Return true if it does
        $headers_match = RETURN-BOOL_Confirm-Headers-in-CSV $csv
    }
    
    #If $headers_match return $true...
    if ($headers_match){
        #Clear old data from table
        $table.Clear()
        #Load data from CSV to the Datagrid
        CSV-to-DataGrid
        #Clean up blank rows in the data grid
        Clean-Up-Blank-Rows-DataGrid
        #Check for Duplicates (Active Directory Names)
        Duplicates-Populate-Mnmemonic-Description-Field
        #Generate Mnemonics
        NonDuplicates-Populate-Mnmemonic-Description-Field
        #Generate Password for all Mnemonics
        Assign-Complex-Password

    #If $headers_match return $false...
    }elseif ($headers_match -eq $false){
        #display message
        write-host "CSV Does not have the proper headers"
    }
})

$CreateADButton.Add_Click({
    Create-AD-Accounts-Button
})

$ExportCSVButton.Add_Click({
    Export-Data-As-Csv
})

#-------------------------------------------------------------#

$Window.ShowDialog() | Out-Null
