<#

***************   This script is provided AS-IS without any warranty to any damage that may occured. It will delete emails, if you are using it it's AT YOUR OWN RISK!  ***********

Version 1.0
Inital release

#>


$str001 = "Contacts creation ver 1.0"
$str002 = "Connect"
$str003 = "Use MFA"
$str004 = "Connection Status:"
$str005 = "Not Connected to Exchnage online"
$str006 = "Connected to Exchnage online"
$str007 = "Display Name:"
$str008 = "External email address:"
$str009 = "Internal email address (if needed):"
$str010 = "Hide From Address Lists"
$str011 = "Wrong creds or no creds entered"
$str012 = "Display name or External email address are missing"
$str013 = "Apply"
$str014 = "Close"


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '600,450'
$Form.text                       = $str001
$Form.TopMost                    = $false

$ConnectButton                   = New-Object system.Windows.Forms.Button
$ConnectButton.text              = $str002
$ConnectButton.Size              = new-object System.Drawing.Size(200,30)
$ConnectButton.location          = New-Object System.Drawing.Point(30,45)
$ConnectButton.Font              = 'Microsoft Sans Serif,10'
$form.Controls.Add($ConnectButton)
$ConnectButton.Add_Click({connect})

$MFA_CheckBox                    = new-object System.Windows.Forms.checkbox
$MFA_CheckBox.Location           = new-object System.Drawing.Size(280,40)
$MFA_CheckBox.Size               = new-object System.Drawing.Size(250,50)
$MFA_CheckBox.Text               = $str003
$MFA_CheckBox.Checked            = $true
$form.Controls.Add($MFA_CheckBox)

$ConnectionStatus                = New-Object system.Windows.Forms.Label
$ConnectionStatus.text           = $str004
$ConnectionStatus.AutoSize       = $true
$ConnectionStatus.location       = New-Object System.Drawing.Point(30,100)
$ConnectionStatus.Size               = new-object System.Drawing.Size(25,10)
$ConnectionStatus.Font           = 'Microsoft Sans Serif,10'
$form.Controls.Add($ConnectionStatus)

$StatusUpdate                    = New-Object system.Windows.Forms.Label
$StatusUpdate.ForeColor          = 'Red'
$StatusUpdate.Text               = $str005
$StatusUpdate.AutoSize           = $true
$StatusUpdate.location           = New-Object System.Drawing.Point(150,100)
$StatusUpdate.Size               = new-object System.Drawing.Size(25,10)
$StatusUpdate.Font               = 'Microsoft Sans Serif,10'
$form.Controls.Add($StatusUpdate)

$DisplayName                      = New-Object system.Windows.Forms.Label
$DisplayName.text                 = $str007
$DisplayName.AutoSize             = $true
$DisplayName.location             = New-Object System.Drawing.Point(30,147)
$DisplayName.Size                = new-object System.Drawing.Size(25,10)
$DisplayName.Font                 = 'Microsoft Sans Serif,10'
$form.Controls.Add($DisplayName)

$DisplayNameTextBox               = New-Object system.Windows.Forms.TextBox
$DisplayNameTextBox.multiline     = $false
$DisplayNameTextBox.location      = New-Object System.Drawing.Point(125,145)
$DisplayNameTextBox.Size          = new-object System.Drawing.Size(250,20)
$DisplayNameTextBox.Font          = 'Microsoft Sans Serif,10'
$form.Controls.Add($DisplayNameTextBox)

$ExternalEmailAddress             = New-Object system.Windows.Forms.Label
$ExternalEmailAddress.text        = $str008
$ExternalEmailAddress.AutoSize    = $true
$ExternalEmailAddress.location    = New-Object System.Drawing.Point(30,200)
$ExternalEmailAddress.Size        = new-object System.Drawing.Size(25,10)
$ExternalEmailAddress.Font        = 'Microsoft Sans Serif,10'
$form.Controls.Add($ExternalEmailAddress)

$ExternalEmailAddressTextBox            = New-Object system.Windows.Forms.TextBox
$ExternalEmailAddressTextBox.multiline  = $false
$ExternalEmailAddressTextBox.location   = New-Object System.Drawing.Point(180,195)
$ExternalEmailAddressTextBox.Size       = new-object System.Drawing.Size(300,30)
$ExternalEmailAddressTextBox.Font       = 'Microsoft Sans Serif,10'
$form.Controls.Add($ExternalEmailAddressTextBox)

$InternalEmailAddress             = New-Object system.Windows.Forms.Label
$InternalEmailAddress.text        = $str009
$InternalEmailAddress.AutoSize    = $true
$InternalEmailAddress.location    = New-Object System.Drawing.Point(30,250)
$InternalEmailAddress.Size        = new-object System.Drawing.Size(25,10)
$InternalEmailAddress.Font        = 'Microsoft Sans Serif,10'
$form.Controls.Add($InternalEmailAddress)

$InternalEmailAddressTextBox            = New-Object system.Windows.Forms.TextBox
$InternalEmailAddressTextBox.multiline  = $false
$InternalEmailAddressTextBox.location   = New-Object System.Drawing.Point(245,245)
$InternalEmailAddressTextBox.Size       = new-object System.Drawing.Size(300,30)
$InternalEmailAddressTextBox.Font       = 'Microsoft Sans Serif,10'
$form.Controls.Add($InternalEmailAddressTextBox)

$HideFromAddressLists_CheckBox          = new-object System.Windows.Forms.checkbox
$HideFromAddressLists_CheckBox.Location = new-object System.Drawing.Size(30,285)
$HideFromAddressLists_CheckBox.Size     = new-object System.Drawing.Size(250,50)
$HideFromAddressLists_CheckBox.Text     = $str010
$form.Controls.Add($HideFromAddressLists_CheckBox)

$ApplyButton                     = New-Object system.Windows.Forms.Button
$ApplyButton.text                = $str013
$ApplyButton.width               = 102
$ApplyButton.height              = 30
$ApplyButton.location            = New-Object System.Drawing.Point(100,350)
$ApplyButton.Font                = 'Microsoft Sans Serif,10'
$form.Controls.Add($ApplyButton)
$ApplyButton.Add_Click({CreateTheContact})

$closeButton                     = New-Object system.Windows.Forms.Button
$closeButton.text                = $str014
$closeButton.width               = 102
$closeButton.height              = 30
$closeButton.location            = New-Object System.Drawing.Point(400,350)
$closeButton.Font                = 'Microsoft Sans Serif,10'
$form.Controls.Add($closeButton)
$closeButton.Add_Click({closeForm})


$MsgBoxError = [System.Windows.Forms.MessageBox]
$MsgBoxNotify = [System.Windows.Forms.MessageBox]
#endregion GUI


function connect ()
{
    if ($MFA_CheckBox.Checked)
    {
        MFA
    }
    else
    {
        Non_MFA
    }  
}

function MFA () 
{
    if (test-path ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1))
    {
        try
        {
            $MFAExchangeModule = ((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1) 
            If ($MFAExchangeModule -eq $null)
            {
                write-host ("Please install Exchange Online MFA Module.")
            }
            else
            {
                ."$MFAExchangeModule"
                $StatusUpdate.Text = "Connecting"
                $StatusUpdate.Forecolor = 'orange'
                Connect-EXOPSSession -WarningAction SilentlyContinue | Out-Null
            }
        }
        catch
        {
            $MsgBoxError::Show("Wrong creds or no creds entered", $str001, "OK", "Error")
        }
        TestConnection
    }
}


function Non_MFA ()
{
    try 
    {
        $StatusUpdate.Text = "Connecting"
        $StatusUpdate.Forecolor = 'orange'
        $URL = "https://ps.outlook.com/powershell"
        $UserCredential = Get-Credential -ErrorAction Continue
        $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URL -Credential $Credentials -Authentication Basic -AllowRedirection -Name "Exchange Online"
        Import-PSSession $EXOSession -DisableNameChecking | Out-Null
    }
    catch
    {
        $MsgBoxError::Show("Wrong creds or no creds entered", $str001, "OK", "Error")
    }
        TestConnection
    }


function CreateTheContact ()
{
    if ($StatusUpdate.Text -eq $str006)
    {
        #Check that Display Name and External Email address existing
        if ((!$DisplayNameTextBox.text) -or (!$ExternalEmailAddressTextBox.text))
        {
            $MsgBoxError::Show("Display name or External email address are missing", $str001, "OK", "Error")
        }
        #Internal email address and hide from address lists
        elseif (($InternalEmailAddressTextBox.text) -and ($HideFromAddressLists_CheckBox.checked))
        {
            New-MailContact -Name $DisplayNameTextBox.text -ExternalEmailAddress $ExternalEmailAddressTextBox.text
            Set-MailContact -Identity $DisplayNameTextBox.text -EmailAddresses $InternalEmailAddressTextBox.text -HiddenFromAddressListsEnabled $True
        }
        #Internal email address only
        elseif (($InternalEmailAddressTextBox.text) -and (!$HideFromAddressLists_CheckBox.checked))
        {
            New-MailContact -Name $DisplayNameTextBox.text -ExternalEmailAddress $ExternalEmailAddressTextBox.text
            Set-MailContact -Identity $DisplayNameTextBox.text -EmailAddresses $InternalEmailAddressTextBox.text
        }
        # No internal email address and hide from address lists
        elseif ((!$InternalEmailAddressTextBox.text) -and ($HideFromAddressLists_CheckBox.checked))
        {
            New-MailContact -Name $DisplayNameTextBox.text -ExternalEmailAddress $ExternalEmailAddressTextBox.text
            Set-MailContact -Identity $DisplayNameTextBox.text -HiddenFromAddressListsEnabled $True         
        }
        #Only External email
        else
        {
            New-MailContact -Name $DisplayNameTextBox.text -ExternalEmailAddress $ExternalEmailAddressTextBox.text
        }
        $MsgBoxNotify::Show('Contcat has been created',$str001,'Ok','Information')
    }
    else
    {
         $MsgBoxError::Show("Please connect to Exchange Online", $str001, "OK", "Error")
         
    }
}

Function TestConnection()
{
    Try
    {
        Get-MailContact -ErrorAction Stop
        $StatusUpdate.ForeColor = 'Green'
        $StatusUpdate.Text = "$str006"
    }
    Catch [System.SystemException]
    {
        $StatusUpdate.ForeColor = 'Red'
        $StatusUpdate.Text = "$str005"
    }
}


function closeForm(){$Form.close()}

[void]$Form.ShowDialog()