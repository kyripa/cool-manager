$global:setGroup = @()
#$global:Credential = Get-Credential



$xHead = @" 
<Window 
 xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="TU Admin Tools" xmlns:d="http://schemas.microsoft.com/expression/blend/2008" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" Height="549.937" Width="787.99">
    <Grid Margin="0,-31,0,0" HorizontalAlignment="Left" Width="777.99" Height="548.937" VerticalAlignment="Top">

"@
$xTabHead = @"
        <TabControl Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="511.453" Margin="0,34.547,0,0" VerticalAlignment="Top" Width="777.007">

"@
$xTabUser = @"
            <TabItem Name="Users" Header="Users">
                <Grid Margin="0,0,2,4">
                    <Button Name="Users_refreshButton" Content="Refresh" HorizontalAlignment="Left" Height="33" Margin="426.003,406,0,0" VerticalAlignment="Top" Width="105"/>
                    <Button Name="Users_kickButton" Content="Kick" HorizontalAlignment="Left" Height="33" Margin="536.003,406,0,0" VerticalAlignment="Top" Width="105"/>
                    <Button Name="Users_sendButton" Content="Send" HorizontalAlignment="Left" Height="33" Margin="646.003,406,0,0" VerticalAlignment="Top" Width="105"/>
                    <GroupBox Header="Users" HorizontalAlignment="Left" Height="411" Margin="22,28,0,0" VerticalAlignment="Top" Width="391">
                        <DataGrid Name="Users_userGrid" Height="376" Margin="10,10,-2,0" VerticalAlignment="Top" HeadersVisibility="Column" Background="#FFF0F0F0" BorderBrush="{x:Null}" CanUserResizeRows="False" AutoGenerateColumns="False" IsReadOnly="True">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Session ID" Binding="{Binding SessionID}" IsReadOnly="False"/>
                                <DataGridTextColumn Header="User" Binding="{Binding User}"/>
                                <DataGridTextColumn Header="Server" Binding="{Binding Server}"/>
                                <DataGridTextColumn Header="SessionState" Binding="{Binding SessionState}"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </GroupBox>
                    <GroupBox Header="Message" HorizontalAlignment="Left" Height="362" Margin="418.003,28,0,0" VerticalAlignment="Top" Width="333">
                        <Grid HorizontalAlignment="Left" Height="338" VerticalAlignment="Top" Width="332" Margin="0,0,-2,0">
                            <TextBox Name="Users_MessageTb" HorizontalAlignment="Left" Height="278" Margin="10,50,0,0" TextWrapping="Wrap" Text="Тело сообщения" VerticalAlignment="Top" Width="294"/>
                            <TextBox Name="Users_ThemeTb" HorizontalAlignment="Left" Height="25" Margin="10,10,0,0" TextWrapping="Wrap" Text="Отдел АСУ" VerticalAlignment="Top" Width="294"/>
                        </Grid>
                    </GroupBox>
                </Grid>
            </TabItem>

"@
$xTabRDS = @"
            <TabItem Name="RDS" Header="RDS">
                <Grid Background="#FFE5E5E5" HorizontalAlignment="Left" Width="764.972" Margin="0,0,0,7">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <DataGrid Name="RDGrid" HorizontalAlignment="Left" Height="423.067" Margin="10,18.426,0,0" VerticalAlignment="Top" Width="267.529" AutoGenerateColumns="False" IsReadOnly="True">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="SessionHost" Binding="{Binding SessionHost}" IsReadOnly="False"/>
                            <DataGridTextColumn Header="NewConnectionAllowed" Binding="{Binding NCA}"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    <StackPanel Name="RR" Margin="322,102,87,236">
                        <RadioButton Name="yR" Content="Yes" IsChecked="True"/>
                        <RadioButton Name="nR" Content="No"/>
                        <RadioButton Name="nuR" Content="No, until reboot"/>
                    </StackPanel>
                    <Button Name="getRDSH" Content="get" HorizontalAlignment="Left" Margin="320.915,169.77,0,0" VerticalAlignment="Top" Width="75"/>
                    <Button Name="setRDSH" Content="set" HorizontalAlignment="Left" Margin="400.794,169.77,0,0" VerticalAlignment="Top" Width="75"/>
                </Grid>
            </TabItem>

"@
$xTabWL = @"
            <TabItem Name="WhiteList" Header="WhiteList">
                <Grid Background="#FFE5E5E5">
                    <Button Name="getWhiteList" Content="get" HorizontalAlignment="Left" Height="25.578" Margin="279.395,171.538,0,0" VerticalAlignment="Top" Width="71.114"/>
                    <Button Name="remWhiteList" Content="rem" HorizontalAlignment="Left" Height="25.578" Margin="432.027,171.538,0,0" VerticalAlignment="Top" Width="71.113"/>
                    <TextBox Name="postLb" HorizontalAlignment="Left" Height="425.022" Margin="18.224,23.471,0,0" VerticalAlignment="Top" Width="245.18" TextWrapping="Wrap" AcceptsReturn="True" HorizontalScrollBarVisibility="Disabled" VerticalScrollBarVisibility="Auto"/>
                    <TextBox Name="postAb"  HorizontalAlignment="Left" Height="23" Margin="279.395,143.538,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="223.924"/>
                    <Button Name="addWhiteList" Content="add" HorizontalAlignment="Left" Height="25.578" Margin="355.896,171.538,0,0" VerticalAlignment="Top" Width="71.116"/>
                </Grid>
            </TabItem>

"@
$xTabAdU = @"
            <TabItem Name="AddUser" Header="addUser">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-7">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button Name="getCSV" Content="Open CSV" HorizontalAlignment="Left" Height="29.493" Margin="10,10,0,0" VerticalAlignment="Top" Width="89"/>
                    <DataGrid Name="addUserGrid" HorizontalAlignment="Left" Height="189" Margin="10,44.493,0,0" VerticalAlignment="Top" Width="751.007"/>
                    <TextBox Name="searchGroupTb" HorizontalAlignment="Left" Height="23" Margin="10,245.493,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="120"/>
                    <DataGrid Name="getGrDataGrid" HorizontalAlignment="Left" Height="139" Margin="10,273.493,0,0" VerticalAlignment="Top" Width="276" HeadersVisibility="None" ColumnWidth="275" IsReadOnly="True"/>
                    <Button Name="GetGr" Content="GetGr" HorizontalAlignment="Left" Margin="135,245.493,0,0" VerticalAlignment="Top" Width="47" Height="23"/>
                    <Button Name="GetUG" Content="GetUG" HorizontalAlignment="Left" Margin="187,245.493,0,0" VerticalAlignment="Top" Width="47" Height="23"/>
                    <DataGrid Name="setGrDataGrid" HorizontalAlignment="Left" Height="105" Margin="337.003,307.493,0,0" VerticalAlignment="Top" Width="260" HeadersVisibility="None" ColumnWidth="260" IsReadOnly="True"/>
                    <Button Name="setGR" Content="&gt;&gt;&gt;" HorizontalAlignment="Left" Margin="291.003,307.66,0,0" VerticalAlignment="Top" Width="41" Height="19.293"/>
                    <Button Name="setOU" Content="&gt;&gt;&gt;" HorizontalAlignment="Left" Margin="291.003,273.66,0,0" VerticalAlignment="Top" Width="41" Height="19.293"/>
                    <Button Name="delGR" Content="&lt;&lt;&lt;" HorizontalAlignment="Left" Margin="291.003,344.41,0,0" VerticalAlignment="Top" Width="41" Height="19.293"/>
                    <TextBlock Name="setOUTb" HorizontalAlignment="Left" Height="18.293" Margin="337.003,274.66,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="424.007" FontSize="10" Text="OU=ТехноЮг,DC=tu,DC=odessa,DC=ua"/>
                    <Button Name="GetOU" Content="GetOU" HorizontalAlignment="Left" Margin="239,245.493,0,0" VerticalAlignment="Top" Width="47" Height="23"/>
                    <DataGrid Name="checkGrid" HorizontalAlignment="Left" Height="55.333" Margin="10,417.493,0,0" VerticalAlignment="Top" Width="751.007" IsReadOnly="True" FontSize="10"/>
                    <Button Name="goButton" Content="GO!" HorizontalAlignment="Left" Margin="602.003,387.033,0,0" VerticalAlignment="Top" Width="90" Height="25.46"/>
                    <Button Name="checkButton" Content="Check" HorizontalAlignment="Left" Margin="602.003,362.073,0,0" VerticalAlignment="Top" Width="90"/>
                    <TextBox HorizontalAlignment="Left" Height="23" Margin="522.503,246.66,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="74.5"/>
                    <Label Content="SIP:" HorizontalAlignment="Left" Margin="488.01,244.867,0,0" VerticalAlignment="Top" Width="29.493" Height="24.793"/>
                </Grid>
            </TabItem>

"@
$xTabDM = @"
            <TabItem Name="Dismiss" Header="Dismiss">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <TextBox Name="searchDisissTb" HorizontalAlignment="Left" Height="23" Margin="7,6.96,0,0" TextWrapping="Wrap" Text="тест" VerticalAlignment="Top" Width="197"/>
                    <DataGrid Name="DismissGrid" HorizontalAlignment="Left" Height="173" Margin="7,34.96,0,0" VerticalAlignment="Top" Width="751.007" IsReadOnly="True"/>
                    <Button Name="searchDismissButton" Content="Search" HorizontalAlignment="Left" Margin="209.003,10,0,0" VerticalAlignment="Top" Width="75"/>
                    <Button Name="dismissButton" Content="Dismiss Selected" HorizontalAlignment="Left" Margin="289.003,10,0,0" VerticalAlignment="Top" Width="114"/>
                    <TextBlock Name="dismissLogTb" HorizontalAlignment="Left" Height="115" Margin="10,225.493,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="628"/>
                </Grid>
            </TabItem>

"@
$xTabTA = @"
            <TabItem Name="TermAccess" Header="TermAccess" HorizontalAlignment="Left" Height="19.96" VerticalAlignment="Top" Width="74.64">
                <Grid Background="#FFE5E5E5">                    
                    <DatePicker Name="TA_DP" HorizontalAlignment="Left" Margin="186.25,39,0,0" VerticalAlignment="Top" Height="23" SelectedDateFormat="Long"/>
                    <TextBox Name="TA_searchGU_Tb" HorizontalAlignment="Left" Height="23" Margin="10,39,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="97"/>
                    <Button Name="TA_getU_B" Content="GetU" HorizontalAlignment="Left" Margin="112,39,0,0" VerticalAlignment="Top" Width="30" Height="23"/>
                    <DataGrid Name="TA_searchGU_DG" HeadersVisibility="None" HorizontalAlignment="Left" Height="100" Margin="10,67,0,0" VerticalAlignment="Top" Width="276.07"/>
                    <Button Name="TA_getG_B" Content="GetG" HorizontalAlignment="Left" Margin="147,39,0,0" VerticalAlignment="Top" Width="30" Height="23"/>
                    <TextBlock Name="TA_setG_Tb" HorizontalAlignment="Left" Margin="339,67,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="388.957" Text="Разрешена удаленка" Height="15.96"/>
                    <DataGrid Name="TA_setGU_DG" HeadersVisibility="None" HorizontalAlignment="Left" Height="79.04" Margin="339,87.96,0,0" VerticalAlignment="Top" Width="218.342"/>
                    <GroupBox Header="NewGroup" HorizontalAlignment="Left" Height="197" Margin="10,203,0,0" VerticalAlignment="Top" Width="276.07">
                        <Grid HorizontalAlignment="Left" Height="174.493" Margin="10,0,-2,-0.453" VerticalAlignment="Top" Width="256.07">
                            <TextBox Name="TA_GroypDisc_Tb" HorizontalAlignment="Left" Height="126.493" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="243.667" Margin="0.333,38,0,0"/>
                            <TextBox Name="TA_GroupName_Tb" HorizontalAlignment="Left" Height="23" Margin="0.333,10,-35,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="243.667"/>
                        </Grid>
                    </GroupBox>
                    <Button Name="TA_setG_B" Content="&gt;&gt;&gt;" HorizontalAlignment="Left" Margin="291.07,67,0,0" VerticalAlignment="Top" Width="42.93"/>
                    <Button Name="TA_setGU_B" Content="&gt;&gt;&gt;" HorizontalAlignment="Left" Margin="291.07,100.666,0,0" VerticalAlignment="Top" Width="42.93"/>
                    <Button Name="TA_delGU_B" Content="&lt;&lt;&lt;" HorizontalAlignment="Left" Margin="291.07,125.626,0,0" VerticalAlignment="Top" Width="42.93"/>
                    <Button Name="TA_go_B" Content="GO!" HorizontalAlignment="Left" Height="31.333" Margin="304,215.493,0,0" VerticalAlignment="Top" Width="70.333"/>
                </Grid>
            </TabItem>

"@
$xTabSettings = @"
            <TabItem Name="Settings" Header="Settings" HorizontalAlignment="Left" Height="19.96" VerticalAlignment="Top" Width="53.9733333333333">
                <Grid Background="#FFE5E5E5">
                    <GroupBox Name="Exchange_settings" Header="Exchange" HorizontalAlignment="Left" Height="129.16" Margin="10,10,0,0" VerticalAlignment="Top" Width="286.333">
                        <Grid Margin="0,0,-2,-1.626">
                            <TextBox Name="Settings_ExConnection_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="http://post.tu.odessa.ua/powershell" VerticalAlignment="Top" Width="256.333" Margin="10,25.96,0,0" FontSize="10"/>
                            <Label Content="Connection:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,0,0,0" FontSize="10"/>
                            <TextBox Name="Settings_ExDefDB_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="MAILBOXTUDB01" VerticalAlignment="Top" Width="256.333" Margin="10,74.92,0,0" FontSize="10"/>
                            <Label Content="Default new user database:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,48.96,0,0" FontSize="10"/>
                        </Grid>
                    </GroupBox>
                    <GroupBox Name="AddDeluser_settings" Header="Add/Del user" HorizontalAlignment="Left" Height="129.16" Margin="301.333,10,0,0" VerticalAlignment="Top" Width="447">
                        <Grid Margin="0,0,-2,-1.626">
                            <TextBox Name="Settings_DefNewUsersOU_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="OU=Пользователи,OU=ТехноЮг,DC=tu,DC=odessa,DC=ua" VerticalAlignment="Top" Width="417" Margin="10,25.96,0,0" FontSize="10"/>
                            <Label Content="Default new users OU:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,0,0,0" FontSize="10"/>
                            <TextBox Name="Settings_DefDismissedUsersOU_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="OU=_Уволенные сотрудники,OU=Пользователи,OU=ТехноЮг,DC=tu,DC=odessa,DC=ua" VerticalAlignment="Top" Width="417" Margin="10,79.92,0,0" FontSize="10"/>
                            <Label Content="Default dismissed users OU:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,53.96,0,0" FontSize="10"/>
                        </Grid>
                    </GroupBox>
                    <GroupBox Name="RDS_settings" Header="RDS" HorizontalAlignment="Left" Height="129.16" Margin="10,144.16,0,0" VerticalAlignment="Top" Width="286.333">
                        <Grid Margin="0,0,-2,-1.626">
                            <TextBox Name="Settings_DefRDCB_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="rdcb.tu.odessa.ua" VerticalAlignment="Top" Width="256.333" Margin="10,25.96,0,0" FontSize="10"/>
                            <Label Content="Connection Broker:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,0,0,0" FontSize="10"/>
                            <TextBox Name="Settings_DefCollection_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="Ферма RDS" VerticalAlignment="Top" Width="256.333" Margin="10,74.92,0,0" FontSize="10"/>
                            <Label Content="Collection:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,48.96,0,0" FontSize="10"/>
                        </Grid>
                    </GroupBox>
                    <Button Name="Settings_save_B" Content="Save" HorizontalAlignment="Left" Margin="10,278.32,0,0" VerticalAlignment="Top" Width="75"/>
                    <GroupBox Name="ActiveDirectory" Header="ActiveDirectory" HorizontalAlignment="Left" Height="129.16" Margin="301.333,144.16,0,0" VerticalAlignment="Top" Width="447">
                        <Grid Margin="0,0,-2,-1.626">
                            <TextBox Name="Settings_DefDC_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="pdc.tu.odessa.ua" VerticalAlignment="Top" Width="417" Margin="10,25.96,0,0" FontSize="10"/>
                            <Label Content="Default DC Server:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,0,0,0" FontSize="10"/>
                            <TextBox Name="Settings_DefTempGroupOU_Tb" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" Text="OU=Временный доступ,OU=Группы доступа,OU=ТехноЮг,DC=tu,DC=odessa,DC=ua" VerticalAlignment="Top" Width="417" Margin="10,79.92,0,0" FontSize="10"/>
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


[xml]$xaml = $xHead + $xTabHead + $xTabUser + $xTabRDS + $xTabWL + $xTabAdU + $xTabDM + $xTabTA + $xTabSettings + $xTabEnd + $xEnd
Add-Type -AssemblyName presentationframework
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$Form=[Windows.Markup.XamlReader]::Load( $reader )
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}

#Default Settings
$global:ExConnection = $Settings_ExConnection_Tb.Text
$global:ExDefDB = $Settings_ExDefDB_Tb.Text
$global:DefNewUsersOU = $Settings_DefNewUsersOU_Tb.Text
$global:DefDismissedUsersOU = $Settings_DefDismissedUsersOU_Tb.Text
$global:DefRDCB = $Settings_DefRDCB_Tb.Text
$global:DefCollection = $Settings_DefCollection_Tb.Text
$global:DefDC = $Settings_DefDC_Tb.Text
$global:DefTempGroupOU = $Settings_DefTempGroupOU_Tb.Text


$setOUTb.Text = $global:DefNewUsersOU
$global:ADmoduleLoad = $false
$global:EXmoduleLoad = $false

#Functions
Function Connect-Mstsc {
    [cmdletbinding(SupportsShouldProcess,DefaultParametersetName='UserPassword')]
    param (
        [Parameter(Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            Position=0)]
        [Alias('CN')]
            [string[]]     $ComputerName,
        [Parameter(ParameterSetName='UserPassword',Mandatory=$true,Position=1)]
        [Alias('U')] 
            [string]       $User,
        [Parameter(ParameterSetName='UserPassword',Mandatory=$true,Position=2)]
        [Alias('P')] 
            [string]       $Password,
        [Parameter(ParameterSetName='Credential',Mandatory=$true,Position=1)]
        [Alias('C')]
            [PSCredential] $Credential,
        [Alias('A')]
            [switch]       $Admin,
        [Alias('MM')]
            [switch]       $MultiMon,
        [Alias('F')]
            [switch]       $FullScreen,
        [Alias('Pu')]
            [switch]       $Public,
        [Alias('W')]
            [int]          $Width,
        [Alias('H')]
            [int]          $Height,
        [Alias('WT')]
            [switch]       $Wait
    )

    begin {
        [string]$MstscArguments = ''
        switch ($true) {
            {$Admin}      {$MstscArguments += '/admin '}
            {$MultiMon}   {$MstscArguments += '/multimon '}
            {$FullScreen} {$MstscArguments += '/f '}
            {$Public}     {$MstscArguments += '/public '}
            {$Width}      {$MstscArguments += "/w:$Width "}
            {$Height}     {$MstscArguments += "/h:$Height "}
        }

        if ($Credential) {
            $User     = $Credential.UserName
            $Password = $Credential.GetNetworkCredential().Password
        }
    }
    process {
        foreach ($Computer in $ComputerName) {
            $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
            $Process = New-Object System.Diagnostics.Process
            
            # Remove the port number for CmdKey otherwise credentials are not entered correctly
            if ($Computer.Contains(':')) {
                $ComputerCmdkey = ($Computer -split ':')[0]
            } else {
                $ComputerCmdkey = $Computer
            }

            $ProcessInfo.FileName    = "$($env:SystemRoot)\system32\cmdkey.exe"
            $ProcessInfo.Arguments   = "/generic:TERMSRV/$ComputerCmdkey /user:$User /pass:$($Password)"
            $ProcessInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
            $Process.StartInfo = $ProcessInfo
            if ($PSCmdlet.ShouldProcess($ComputerCmdkey,'Adding credentials to store')) {
                [void]$Process.Start()
            }

            $ProcessInfo.FileName    = "$($env:SystemRoot)\system32\mstsc.exe"
            $ProcessInfo.Arguments   = "$MstscArguments /v $Computer"
            $ProcessInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Normal
            $Process.StartInfo       = $ProcessInfo
            if ($PSCmdlet.ShouldProcess($Computer,'Connecting mstsc')) {
                [void]$Process.Start()
                if ($Wait) {
                    $null = $Process.WaitForExit()
                }       
            }
        }
    }
}

function connectAD{
        param(
            $dc
        )
        $global:ADmoduleLoad = $true
        $session = New-PSSession -ComputerName $dc -Name "PDC"
        Invoke-Command $session -Scriptblock { Import-Module ActiveDirectory }
        Import-PSSession $Session -Module ActiveDirectory           
}

function connectExchange{
         param(
            $exchange
         )
         $global:EXmoduleLoad = $true
         $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchange -Name "Exchange"
         Import-PSSession $Session -CommandName *ContentFilterConfig, New-Mailbox
}

#Settings tab begin
$Settings_save_B.add_Click({
    $global:ExConnection = $Settings_ExConnection_Tb.Text
    $global:ExDefDB = $Settings_ExDefDB_Tb.Text
    $global:DefNewUsersOU = $Settings_DefNewUsersOU_Tb.Text
    $global:DefDismissedUsersOU = $Settings_DefDismissedUsersOU_Tb.Text
    $global:DefRDCB = $Settings_DefRDCB_Tb.Text
    $global:DefCollection = $Settings_DefCollection_Tb.Text
    $global:DefDC = $Settings_DefDC_Tb.Text
    $global:DefTempGroupOU = $Settings_DefTempGroupOU_Tb.Text

    $global:ADmoduleLoad = $false
    $global:EXmoduleLoad = $false


    $setOUTb.Text = $global:DefNewUsersOU

    Remove-PSSession -Name "Exchange" -ErrorAction Ignore
    Remove-PSSession -Name "PDC" -ErrorAction Ignore



})
#Settings tab end
#TermAccess tab begin
$TA_DP.add_SelectedDateChanged({
  $TA_GroypDisc_Tb.Text = "с $((Get-Date).ToString('dd.MM.yyyy')) по $($TA_DP.SelectedDate.AddHours(23).ToString('dd.MM.yyyy-HH:mm'))"  

})
$TA_getG_B.add_Click({
    if(!$global:ADmoduleLoad){
        connectAD -dc $DefDC           
    }


    $Uname = $TA_searchGU_Tb.text
    $g = Get-ADGroup -filter "name -like `"*$Uname*`"" | Select-Object SamAccountName
    $TA_setG_B.IsEnabled = $true

   
    $TA_searchGU_DG.ItemsSource = @($g)
})
$TA_getU_B.add_Click({
    if(!$global:ADmoduleLoad){
        connectAD -dc $DefDC           
    }


    $Uname = $TA_searchGU_Tb.text
    $g = Get-ADuser -filter "name -like `"*$Uname*`"" | Select-Object SamAccountName
    
    $TA_setG_B.IsEnabled = $false

   
    $TA_searchGU_DG.ItemsSource = @($g)
})
$TA_setGU_B.add_Click({
    [psobject[]]$global:setGroup += $TA_searchGU_DG.SelectedItems
    $global:setGroup = $global:setGroup | select SamAccountName -Unique
    $TA_setGU_DG.ItemsSource = @($global:setGroup)
})
$TA_delGU_B.add_Click({
    $grR = $TA_setGU_DG.SelectedItems
    $grAll = $TA_setGU_DG.Items
    $global:setGroup = Compare-Object -ReferenceObject $grR -DifferenceObject $grAll -PassThru -Property SamAccountName
    
    $global:setGroup = $global:setGroup | select SamAccountName -Unique
    #$global:setGroup = $global:GR
    
    $TA_setGU_DG.ItemsSource = @($global:setGroup)
    
})
$TA_setG_B.add_Click({
    [psobject[]]$setGroup = $TA_searchGU_DG.SelectedItems
    if($setGroup.Length -ge 1){
        $TA_setG_Tb.text = [string]$setGroup[0].SamAccountName
    }     
})
$TA_go_B.add_Click({
    if(($TA_GroupName_Tb.text -ne "") -and ($TA_GroypDisc_Tb.Text -ne "")){
        if(!$global:ADmoduleLoad){
            connectAD -dc $DefDC           
        }
        $groupName = $TA_GroupName_Tb.Text
        $groupDisc = $TA_GroypDisc_Tb.Text
        $rootGroup = $TA_setG_Tb.Text
        $groupeMembers = $TA_setGU_DG.Items

        $dayX = $TA_DP.SelectedDate.AddHours(23)
        [int]$TTL =$($dayX - $(Get-Date)).TotalSeconds

        $OU = [adsi]"LDAP://OU=Временный доступ,OU=Группы доступа,OU=ТехноЮг,DC=tu,DC=odessa,DC=ua"
        $Group = $OU.Create("group","cn=$groupName")
        $Group.PutEx(2,"objectClass",@("dynamicObject","group"))
        $Group.Put("entryTTL","$TTL")
        $Group.Put("sAMAccountName","$groupName")
        $Group.Put("description","$groupDisc")
        $Group.SetInfo()
        #
        Start-Sleep -Seconds 20

        Add-ADGroupMember -Identity $rootGroup -Members $groupName
        foreach($member in $groupeMembers){
            try{
            Add-ADGroupMember -Identity $groupName -Members $member.SamAccountName  
            }catch{
                $msgBoxInput = [System.Windows.Forms.MessageBox]::Show($member.SamAccountName + ' cant add in group', 'Logoff', 'ok', 'Error')
            }     
        }
        #>


    } else{
        $msgBoxInput = [System.Windows.Forms.MessageBox]::Show('GroupName or GroupDiscription does not exist', 'Logoff', 'ok', 'Error')
    }
#> 
})
#TermAccess tab end

#AddUser tab begin
$getCSV.add_Click({

    $OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "Excell (*.xls)|*.xls|All files (*.*)|*.*";
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
        $i = Import-CSV -Path $csv -delimiter ";" -Encoding Default  -WarningAction Ignore
        $addUserGrid.itemsSource = @($i)
    }
})
$getGr.add_Click({
    if(!$global:ADmoduleLoad){
        connectAD -dc $DefDC           
    }


    $Dname = $searchGroupTb.text
    $g = Get-ADGroup -filter "name -like `"*$Dname*`"" | Select-Object name
    
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
     $userSIP = $sipTb.text
     $checkUsers = @()

foreach($user in $userCSV) {  
            New-Mailbox -Name $user.name `
                        -Displayname $user.name `
                        -Lastname $user.”lastName” `
                        -FirstName $user.”firstName” `
                        -OrganizationalUnit $userOU `
                        -Database $global:ExDefDB `
                        -UserPrincipalName ($user.”login” + "@tu.odessa.ua")`
                        -Password (ConvertTo-SecureString $user."pass" -AsPlainText -Force)
            Start-Sleep -Seconds 1
            Set-ADUser $user."login" -Company "ТехноЮг" #TODO custom company
            Set-ADUser $user."login" -OfficePhone $userSIP -WarningAction SilentlyContinue
            Set-ADUser $user."login" -Department $user."department"
            Set-ADUser $user."login" -Title $user."title"
            Set-ADUser $user."login" -Office $user."office"            
            Set-ADUser $user."login" -PasswordNeverExpires $true -CannotChangePassword $true -ChangePasswordAtLogon $false -Enabled $true
            foreach($gr in $userGR.name){
                Add-ADGroupMember -Identity $gr -Members $user."login"
            }
            Start-Sleep -Seconds 1            
            $checkUsers += get-ADUser $user."login" | Select-Object Enabled, Name, DistinguishedName
            
}
$checkGrid.itemsSource = @($checkUsers)


})
<#
$checkButton.add_Click({
    $userCSV = $addUserGrid.SelectedItems

    foreach($user in $userCSV){
      Connect-Mstsc -ComputerName 192.168.81.120 -User "tu\$($user.login)" -Password "$($user.pass)"
    }


})
#>
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

    foreach($user in $checkUsers){
        if($user.enabled -eq "true"){
            $msgBoxInput = [System.Windows.Forms.MessageBox]::Show('You really want connect mstsc to '+ $user.name +'?', 'Connect', 'YesNo', 'Warning')
                switch ($msgBoxInput) {

                    'Yes' {
                        Connect-Mstsc -ComputerName 192.168.81.120 -User "tu\$($user.login)" -Password "$($user.pass)"
                        }

                    'No' {
                        #Do nothing
                        }
                    }
                }
            }



})
#>
#Adduser tab end
 
#Dismiss tab begin
$searchDismissButton.Add_Click({
    
    if(!$global:ADmoduleLoad){
        connectAD -dc $DefDC           
    }

    $Dname = "*" + $searchDisissTb.text + "*"
    $user = Get-ADUser -filter "name -like `"$Dname`"" | Select-Object Enabled, Name, DistinguishedName  
           
    $DismissGrid.itemsSource = @($user)

})
$dismissButton.Add_Click({
    $users = $DismissGrid.SelectedItems
    
    if(!$global:ADmoduleLoad){
        connectAD -dc $DefDC           
    } 
    $dismissLogTb.Text = ""
    foreach($user in $users){          
                $u = Move-ADObject $user.DistinguishedName -TargetPath $global:DefDismissedUsersOU -PassThru
                $b = $u.tostring().split(",")
                $dismissLogTb.Text += $b[0] + " moved to " + $b[1] + "`n"
    }

})
#Dismiss tab end
 
#RDS tab begin
$getRDSH.Add_Click( {
        $rdhosts = Get-RDSessionHost -ConnectionBroker $global:DefRDCB -CollectionName $global:DefCollection
        $item = @()
        foreach ($RDhost in $rdhosts) {
            $row = new-object PSObject -Property @{
                SessionHost = $RDhost.SessionHost;                
                NCA = $RDhost.NewConnectionAllowed;
            }
            $item += $row           
        }

        $RDGrid.ItemsSource = @($item)
       
 })
$setRDSH.Add_Click( {

    if($yR.IsChecked){   
        $NCA= "Yes"
    } else{    
        if($nR.IsChecked){
            $NCA= "No"
        }else{
            if($nuR.IsChecked){
                $NCA= "NotUntilReboot"
                    }
                }
    }
    $RDHosts = $RDGrid.SelectedItems
    foreach($RDhost in $RDHosts){
        Set-RDSessionHost -SessionHost $RDhost.SessionHost -NewConnectionAllowed $NCA -ConnectionBroker $global:DefRDCB   
    }        
 })
#RDS tab end 

#WhiteList tab begin
function getBypassedSenderDomains{
    $postLb.Text = ""
    $item =  (get-contentfilterconfig).bypassedSenderDomains
    foreach($i in $item){
        $postLb.text += $i + "`n"
    }

}
$getWhiteList.Add_Click( {
    if(!$global:EXmoduleLoad){
        connectExchange -exchange $ExConnection
     }
    getBypassedSenderDomains} )
$addWhiteList.Add_Click({
    if(!$global:EXmoduleLoad){
        connectExchange -exchange $ExConnection
     }
    
    $item = $postAb.text
        
    if($item -ne ""){   
        Set-ContentFilterConfig –BypassedSenderDomains @{Add=$item}
    }
    getBypassedSenderDomains    
})
$remWhiteList.Add_Click({
    if(!$global:EXmoduleLoad){
        connectExchange -exchange $ExConnection
     }

    $item = $postAb.text 
    if($item -ne ""){    
    Set-ContentFilterConfig –BypassedSenderDomains @{remove=$item}
    }
    getBypassedSenderDomains    
})
#WhiteList tab end 

#users tab begin
$Users_refreshButton.Add_Click( { 
        $Usersessions = Get-RDUserSession -ConnectionBroker $global:DefRDCB -Verbose | Select-Object UnifiedSessionID, HostServer, CollectionName, Username, SessionState
        $item = @()
        foreach ($UserSession in $Usersessions) {
            $row = new-object PSObject -Property @{
                SessionID = $Usersession.UnifiedSessionID;                
                User = $Usersession.Username;
                Server = $UserSession.HostServer.Split(".")[0];
                SessionState = $Usersession.SessionState
            }
            $item += $row           
        }

        $Users_userGrid.ItemsSource = @($item)
        })
$Users_kickButton.Add_Click( {
        $chosenOne = $Users_userGrid.SelectedItems

        if($chosenOne.count -ne 0){
        $msgBoxInput = [System.Windows.Forms.MessageBox]::Show('You really want Logoff ' + $chosenOne.count + ' users?', 'Logoff', 'YesNo', 'Error')

            switch ($msgBoxInput) {

                'Yes' {
                    foreach($one in $chosenOne){
                        Invoke-RDUserLogoff -HostServer $one.Server -UnifiedSessionID $one.SessionID -Force
                    }
                }

                'No' {
                    #Do nothing
                    }
                }
            }
        })
$Users_sendButton.Add_Click( {             
        $chosenOne = $Users_userGrid.SelectedItems
         

        $theme = $Users_ThemeTb.Text
        $message = $Users_MessageTb.Text
        
        if($chosenOne.count -ne 0){
        $msgBoxInput = [System.Windows.Forms.MessageBox]::Show('You really want send message ' + $chosenOne.count + ' users?', 'Send message', 'YesNo', 'info')

            switch ($msgBoxInput) {

                'Yes' {
                    foreach($one in $chosenOne){                       
                        Send-RDUserMessage -HostServer $one.Server -UnifiedSessionID $one.SessionID -MessageTitle $theme -MessageBody $message
                    }
                }

                'No' {
                    #Do nothing
                    }
                }
            }

        })
#user tab end 
  
$Form.Add_Closing({
    try{
        Remove-PSSession -Name "Exchange" -ErrorAction Ignore
        Remove-PSSession -Name "PDC" -ErrorAction Ignore
    }catch{}
    
})

$Form.ShowDialog() | out-null
