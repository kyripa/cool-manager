﻿

$global:setGroup = @()


$xHead = @" 
<Window 
 xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="Add Users" xmlns:d="http://schemas.microsoft.com/expression/blend/2008" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" Height="643.731" Width="1040.918">
    <Grid Margin="0,-31,0,0" HorizontalAlignment="Left" Width="1037.241" Height="648" VerticalAlignment="Top">

"@
$xTabHead = @"
             <TabControl HorizontalAlignment="Left" Height="605.561" Margin="0,34.547,0,0" VerticalAlignment="Top" Width="1034.618">

"@

$xTabAdU = @"
            <TabItem Name="AddUser" Header="addUser">
            <Grid Background="#FFE5E5E5" Margin="0,0,0,-7">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
        </Grid.ColumnDefinitions>
        <Button Name="getCSV" Content="Open CSV" HorizontalAlignment="Left" Height="29.493" Margin="10,10,0,0" VerticalAlignment="Top" Width="89"/>
        <DataGrid Name="addUserGrid" HorizontalAlignment="Left" Height="189" Margin="10,44.493,0,0" VerticalAlignment="Top" Width="1002.796"/>
        <TextBox Name="searchGroupTb" HorizontalAlignment="Left" Height="23" Margin="10,245.493,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="243.302"/>
        <DataGrid Name="getGrDataGrid" HorizontalAlignment="Left" Height="139" Margin="10,273.493,0,0" VerticalAlignment="Top" Width="413.002" HeadersVisibility="None" ColumnWidth="275" IsReadOnly="True"/>
        <Button Name="GetGr" Content="GetGr" HorizontalAlignment="Left" Margin="272.002,245.493,0,0" VerticalAlignment="Top" Width="47" Height="23"/>
        <Button Name="GetUG" Content="GetUG" HorizontalAlignment="Left" Margin="324.002,245.493,0,0" VerticalAlignment="Top" Width="47" Height="23"/>
        <DataGrid Name="setGrDataGrid" HorizontalAlignment="Left" Height="105" Margin="504.567,306.493,0,0" VerticalAlignment="Top" Width="366.44" HeadersVisibility="None" ColumnWidth="260" IsReadOnly="True"/>
        <Button Name="setGR" Content="&gt;&gt;&gt;" HorizontalAlignment="Left" Margin="448.577,306.493,0,0" VerticalAlignment="Top" Width="41" Height="19.293"/>
        <Button Name="setOU" Content="&gt;&gt;&gt;" HorizontalAlignment="Left" Margin="448.577,272.493,0,0" VerticalAlignment="Top" Width="41" Height="19.293"/>
        <Button Name="delGR" Content="&lt;&lt;&lt;" HorizontalAlignment="Left" Margin="448.577,343.243,0,0" VerticalAlignment="Top" Width="41" Height="19.293"/>
        <TextBlock Name="setOUTb" HorizontalAlignment="Left" Height="18.293" Margin="494.577,273.493,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="424.007" FontSize="10" Text=""/>
        <Button Name="GetOU" Content="GetOU" HorizontalAlignment="Left" Margin="376.002,245.493,0,0" VerticalAlignment="Top" Width="47" Height="23"/>
        <DataGrid Name="checkGrid" HorizontalAlignment="Left" Height="153.342" Margin="10,417.493,0,0" VerticalAlignment="Top" Width="1002.796" IsReadOnly="True" FontSize="10"/>
        <Button Name="goButton" Content="GO!" HorizontalAlignment="Left" Margin="890.762,381.764,0,0" VerticalAlignment="Top" Width="90" Height="25.46"/>
        <Button Name="checkButton" Content="Check" HorizontalAlignment="Left" Margin="890.762,356.803,0,0" VerticalAlignment="Top" Width="90"/>
    </Grid>
            </TabItem>

"@
$xTabSettings = @"
        <TabItem Name="Settings" Header="Settings" HorizontalAlignment="Left" Height="19.96" VerticalAlignment="Top" Width="53.9733333333333">
                <Grid Background="#FFE5E5E5">
                    <GroupBox Name="AddDeluser_settings" Header="Add/Del user" HorizontalAlignment="Left" Height="129.16" Margin="10,10,0,0" VerticalAlignment="Top" Width="447">
                        <Grid Margin="0,0,-2,-1.626">
                            <TextBox Name="Settings_DefNewUsersOU_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="CN=Users,DC=mirs,DC=itd" VerticalAlignment="Top" Width="417" Margin="10,25.96,0,0" FontSize="10"/>
                            <Label Content="Default new users OU:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,0,0,0" FontSize="10"/>
                            <TextBox Name="Settings_DefDismissedUsersOU_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="CN=Users,DC=mirs,DC=itd" VerticalAlignment="Top" Width="417" Margin="10,79.92,0,0" FontSize="10"/>
                            <Label Content="Default dismissed users OU:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,53.96,0,0" FontSize="10"/>
                        </Grid>
                    </GroupBox>
                    <Button Name="Settings_save_B" Content="Save" HorizontalAlignment="Left" Margin="10,278.32,0,0" VerticalAlignment="Top" Width="75"/>
                    <GroupBox Name="ActiveDirectory" Header="ActiveDirectory" HorizontalAlignment="Left" Height="129.16" Margin="10,144.16,0,0" VerticalAlignment="Top" Width="447">
                        <Grid Margin="0,0,-2,-1.626">
                            <TextBox Name="Settings_DefDC_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="dc.mirs.itd" VerticalAlignment="Top" Width="417" Margin="10,25.96,0,0" FontSize="10"/>
                            <Label Content="Default DC Server:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,0,0,0" FontSize="10"/>
                            <TextBox Name="Settings_DefTempGroupOU_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="CN=Users,DC=mirs,DC=itd" VerticalAlignment="Top" Width="417" Margin="10,79.92,0,0" FontSize="10"/>
                            <Label Content="Default temp group OU:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,53.96,0,0" FontSize="10"/>
                        </Grid>
                    </GroupBox>
                </Grid>
            </TabItem>

"@
$xTabEnd = @"
</TabControl>

"@
$xEnd = @"
    </Grid>
</Window>

"@


[xml]$xaml = $xHead + $xTabHead + $xTabAdU + $xTabSettings + $xTabEnd + $xEnd
Add-Type -AssemblyName presentationframework
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Form=[Windows.Markup.XamlReader]::Load( $reader )
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}

#Default Settings
$global:DefNewUsersOU = $Settings_DefNewUsersOU_Tb.Text
$global:DefDismissedUsersOU = $Settings_DefDismissedUsersOU_Tb.Text

$global:DefDC = $Settings_DefDC_Tb.Text
$global:DefTempGroupOU = $Settings_DefTempGroupOU_Tb.Text

$setOUTb.Text = $global:DefNewUsersOU

$global:ADmoduleLoad = $false
$global:EXmoduleLoad = $false

#Functions
function connectAD{
        param(
            $dc
        )
        $global:ADmoduleLoad = $true
        $session = New-PSSession -ComputerName $dc -Name "PDC"
        Invoke-Command $session -Scriptblock { Import-Module ActiveDirectory }
        Import-PSSession $Session -Module ActiveDirectory           
}


#Settings tab begin
$Settings_save_B.add_Click({

    $global:DefNewUsersOU = $Settings_DefNewUsersOU_Tb.Text
    $global:DefDismissedUsersOU = $Settings_DefDismissedUsersOU_Tb.Text

    $global:DefDC = $Settings_DefDC_Tb.Text
    $global:DefTempGroupOU = $Settings_DefTempGroupOU_Tb.Text


    $global:ADmoduleLoad = $false

    $setOUTb.Text = $global:DefNewUsersOU

    Remove-PSSession -Name "PDC" -ErrorAction Ignore

})
#Settings tab end


#AddUser tab begin
$getCSV.add_Click({

    $OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "Excell (*.xlsm)|*.xlsm|All files (*.*)|*.*";
    $OpenFileDialog.ShowDialog()

    
    if ($OpenFileDialog.CheckPathExists){ 
        $filePath = $OpenFileDialog.FileName
        $xl = new-object -com Excel.Application 
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.Workbooks.open($filePath)
        $ws = $xl.sheets.item('Лист1')
        $csv = $env:TEMP + "temp.csv" #TODO custom excell
        $ws.SaveAs($csv, 6)
    
        $xl.Workbooks.Close()
        $xl.Quit()
        $i = @()
        $i = Import-CSV -Path $csv -delimiter "," -Encoding Default  -WarningAction Ignore
        $addUserGrid.itemsSource = @($i)
    }
})
$getGr.add_Click({
    if(!$global:ADmoduleLoad){
        connectAD -dc $DefDC           
    }

        $Dname = $searchGroupTb.text

    
    $g = Get-ADGroup -filter "name -like `"$Dname*`"" | Select-Object name
    
    $delGR.IsEnabled = $true
    $setGR.IsEnabled = $true 
    $setOU.IsEnabled = $false   
    $getGrDataGrid.ItemsSource = @($g)

})
$getUG.add_Click({
    if(!$global:ADmoduleLoad){
        connectAD -dc $DefDC           
    }


    $Dname = $searchGroupTb.text

    $us = @()
    $us = Get-ADUser -filter "name -like `"*$Dname*`""
    $g = @()
    if($us -ne $null){
        $g = Get-ADPrincipalGroupMembership $us[0].SamAccountName | Select-Object name
    }
    $delGR.IsEnabled = $true
    $setGR.IsEnabled = $true 
    $setOU.IsEnabled = $false
    $getGrDataGrid.ItemsSource = @($g)

})
$setGR.add_Click({
    [psobject[]]$global:setGroup += $getGrDataGrid.SelectedItems
    $global:setGroup = $global:setGroup | select name -Unique
    $setGrDataGrid.ItemsSource = @($global:setGroup)
})
$delGR.add_Click({
    $grR = $setGrDataGrid.SelectedItems
    $grAll = $setGrDataGrid.Items
    $global:setGroup = Compare-Object -ReferenceObject $grR -DifferenceObject $grAll -PassThru -Property name
    
    $global:setGroup = $global:setGroup | select name -Unique
    #$global:setGroup = $global:GR
    
    $setGrDataGrid.ItemsSource = @($global:setGroup)
    
})
$getOU.add_Click({
    if(!$global:ADmoduleLoad){
        connectAD -dc $DefDC           
    }
       
    $ou = @()
    $ou = Get-ADOrganizationalUnit -SearchBase $global:DefNewUsersOU -Filter {ou -like "*"} | Select-Object DistinguishedName   #TODO custom search base
    $getGrDataGrid.ItemsSource = @($ou)
    $delGR.IsEnabled = $false
    $setGR.IsEnabled = $false 
    $setOU.IsEnabled = $true
})
$setOU.add_Click({   
    #$ou = @()
    $ou = $getGrDataGrid.SelectedItems
    $setOUTb.text = $ou[0].DistinguishedName
})


$goButton.add_Click({
    if(!$global:ADmoduleLoad){
        connectAD -dc $DefDC           
    }
    if(!$global:EXmoduleLoad){
        connectExchange -exchange $ExConnection
     }

     $userCSV = $addUserGrid.SelectedItems
     
     
     $userGR = $setGrDataGrid.Items  
     
     $userOU = $setOUTb.text 
     $checkUsers = @()

foreach($user in $userCSV) {  
            New-ADUser  -Name $user.login `
                        -DisplayName $user.name `
                        -GivenName $user.firstname `
                        -Surname $user.lastname `
                        -AccountPassword (ConvertTo-SecureString $user.pass -AsPlainText -Force) `
                        -Path "$userOU" `
                        -UserPrincipalName ($user.login + '@mirs.itd')
            
            Start-Sleep -Seconds 1
            
            Set-ADUser $user."login" -MobilePhone $user.korpTel -WarningAction SilentlyContinue
            Set-ADUser $user."login" -HomePhone $user.otherTel -WarningAction SilentlyContinue
            Set-ADUser $user."login" -OfficePhone $user.officeTel -WarningAction SilentlyContinue

            Set-ADUser $user."login" -EmailAddress $user.email

            Set-ADUser $user."login" -Company $user.department 
            Set-ADUser $user."login" -Department $user.otdel
            Set-ADUser $user."login" -Title $user.title
            Set-ADUser $user."login" -Office $user.city 
                      
            Set-ADUser $user."login" -PasswordNeverExpires $true -CannotChangePassword $false -ChangePasswordAtLogon $false -Enabled $true
            
            foreach($gr in $userGR.name){
                Add-ADGroupMember -Identity $gr -Members $user."login"
            }
            Start-Sleep -Seconds 1            
            $checkUsers += get-ADUser $user."login" | Select-Object Enabled, Name, DistinguishedName
            
}
$checkGrid.itemsSource = @($checkUsers)


})

$checkButton.add_Click({
    if(!$global:ADmoduleLoad){
        connectAD -dc $DefDC           
    }
    $users = $addUserGrid.SelectedItems
    $checkUsers =@()
    foreach($user in $users) {

        try{
                $U = get-ADUser $user."login" -ErrorAction Stop
                $enabled = $U.enabled
                $U = new-object PSObject -Property @{
                    Enabled = $enabled;                
                    Name = $user.name;
                    login = $user.login;
                    pass = $user.pass}
                }catch{                
                $U = new-object PSObject -Property @{
                    Enabled = "does not exist";                
                    Name = $user.name;
                }
                
                }
        $checkUsers += $U | Select-Object Enabled, Name, login, pass
    }

  
    $checkGrid.itemsSource = @($checkUsers)
})
#>
#Adduser tab end

#  
$Form.Add_Closing({
    try{
        Remove-PSSession -Name "PDC" -ErrorAction Ignore
    }catch{}
    
})
#>
$Form.ShowDialog() | out-null