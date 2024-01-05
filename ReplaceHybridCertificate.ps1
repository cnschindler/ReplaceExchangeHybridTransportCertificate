# Define Logwriter function and variables
#
# Set the Directory for Logging to the directory where the scripts was executed from
$BasePath = Split-Path -Parent $MyInvocation.MyCommand.Path
# Prefix to use for the Logfilename
[string]$LogFileNamePrefix = "ReplaceHybridCertificate"
# Build the Full Path for the Logfile including Date/Time
[string]$LogfileFullPath = Join-Path -Path $BasePath ($LogFileNamePrefix + "_{0:yyyyMMdd-HHmmss}.log" -f [DateTime]::Now)
$Script:NoLogging


function Write-LogFile
{
    # Logging function, used for progress and error logging...
    # Uses the globally (script scoped) configured variables 'LogfileFullPath' to identify the logfile and 'NoLogging' to disable it.
    #
    [CmdLetBinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$LogPrefix,
        [System.Management.Automation.ErrorRecord]$ErrorInfo = $null
    )

    # Prefix the string to write with the current Date and Time, add error message if present...
    if ($ErrorInfo)
    {
        $logLine = "{0:d.M.y H:mm:ss} : ERROR {1}: {2} Error: {3}" -f [DateTime]::Now, $LogPrefix, $Message, $ErrorInfo.Exception.Message
    }

    else
    {
        $logLine = "{0:d.M.y H:mm:ss} : INFO {1}: {2}" -f [DateTime]::Now, $LogPrefix, $Message
    }

    if (-not $NoLogging)
    {
        # Create the Script:Logfile and folder structure if it doesn't exist
        if (-not (Test-Path $LogfileFullPath -PathType Leaf))
        {
            New-Item -ItemType File -Path $LogfileFullPath -Force -Confirm:$false -WhatIf:$false | Out-Null
            Add-Content -Value "Logging started." -Path $LogfileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        }

        # Write to the Script:Logfile
        Add-Content -Value $logLine -Path $LogfileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        Write-Verbose $logLine
    }
    else
    {
        Write-Host $logLine
    }
}

function ConnectExchange
{
    # Check if a connection to an exchange server exists and connect if necessary...
    if (-NOT (Get-PSSession | Where-Object ConfigurationName -EQ "Microsoft.Exchange"))
    {
        $LogPrefix = "ConnectExchange"

        # Test if Exchange Management Shell Module is installed - if not, exit the script
        $EMSModuleFile = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath + "bin\RemoteExchange.ps1"
        
        # If the EMS Module wasn't found
        if (-Not (Test-Path $EMSModuleFile))
        {
            # Write Error end exit the script
            $ErrorMessage = "Exchange Management Shell Module not found on this computer. Please run this script on a computer with Exchange Management Tools installed!"
            Write-LogFile -LogPrefix $LogPrefix -Message $ErrorMessage
            Write-Host -ForegroundColor Red -Message $ErrorMessage
            Exit
        }

        # Load Exchange Management Shell
        try
        {
            . $($EMSModuleFile) -ErrorAction Stop | Out-Null
            Write-LogFile -LogPrefix $LogPrefix -Message "Successfully loaded Exchange Management Shell Module."
        }

        catch
        {
            Write-LogFile -LogPrefix $LogPrefix -Message "Unable to load Exchange Management Shell Module." -ErrorInfo $_
        }

        # Connect to Exchange Server
        try
        {
            Connect-ExchangeServer -auto -ClientApplication:ManagementShell -ErrorAction Stop | Out-Null
            Write-LogFile -LogPrefix $LogPrefix -Message "Successfully connected to Exchange Server."
        }

        catch
        {
            Write-LogFile -LogPrefix $LogPrefix -Message "Unable to connect to Exchange Server." -ErrorInfo $_
        }
    }
}

function Get-ExchangeServers
{
    try
    {
        Write-LogFile -Message "Loading list of Exchange Servers"
        $ExchangeServers = Get-ExchangeServer -Status -ErrorAction Stop
        Write-LogFile -Message "Loaded list of $($ExchangeServers.Count) Exchange Servers"
    }
    
    catch
    {
        Write-LogFile -Message "Failed to load list of Exchange Servers" -ErrorInfo $_
    }

    Return $ExchangeServers
}

function Get-CommonExchangeServerCertificates
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [System.Collections.ArrayList]
        $ExchangeServers
    )

    $AllCerts = New-Object System.Collections.ArrayList 

    Write-LogFile -Message "Building Array of certificates"

    foreach ($Server in $ExchangeServers)
    {
        try
        {
            Write-LogFile -Message "Loading certificates from Server $($Server)"
            $certs = Get-ExchangeCertificate -Server $Server | Where-Object { $_.Subject -ne "CN=Microsoft Exchange Server Auth Certificate" -and $_.Services -match "SMTP" } -ErrorAction Stop
            Write-LogFile -Message "Loaded $($certs.Count) certificates from Server $($Server)"
            foreach ($cert in $certs)
            {
                Write-LogFile -Message "Found certificate $($cert.Subject) on server $($Server)"
            }
        }

        catch
        {
            Write-LogFile -Message "Failed to load certificates from Server $($Server)" -ErrorInfo $_
            $CertificateLoadError = $true
            Break
        }

        Write-LogFile -Message "Adding $($certs.Count) certificates from server $($Server) to Array of certificates"
        foreach ($cert in $certs)
        {
            $ArrayEntry = ($cert.SubjectName.Name.Replace("CN=", "") + ", Expires:" + $cert.NotAfter + ", " + $cert.Thumbprint)
            $AllCerts.Add($ArrayEntry) | Out-Null
        }
    }

    Write-LogFile -Message "Array of certificates contains $($AllCerts.Count) certificates"

    If ($CertificateLoadError)
    {
        $CommonCertificates = $null
        Write-LogFile -Message "Unable to build list of suitable certificates if loading certificates from on or more servers failed"
    }

    elseif ($ExchangeServers.Count -eq 1)
    {
        Write-LogFile -Message "Finding common certificates across Exchange Servers"
        $CommonCertificates = $AllCerts

        if ($CommonCertificates -gt 0)
        {
            Write-LogFile -Message "Found $($CommonCertificates.count) certificates suitable for use as hybrid certificate"
        }
    
        else
        {
            Write-LogFile -Message "No suitable certificates found. Make sure all selected server have the same certificate installed and that those are enabled for SMTP"
        }
    }

    else
    {
        Write-LogFile -Message "Finding common certificates across Exchange Servers"

        $CommonCertificates = $AllCerts | Group-Object | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name

        if ($CommonCertificates -gt 0)
        {
            Write-LogFile -Message "Found $($CommonCertificates.count) certificates suitable for use as hybrid certificate"
        }
    
        else
        {
            Write-LogFile -Message "No suitable certificates found. Make sure all selected server have the same certificate installed and that those are enabled for SMTP"
        }
    }

    Return $CommonCertificates
}

Function Get-CertificateString
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $Thumbprint,
        [Parameter(Mandatory = $true)]
        [string]
        $Server
    )

    
    try
    {
        Write-LogFile -Message "Loading certificate with thumbprint $($Thumbprint) from server $($Server)"
        $TLSCert = Get-ExchangeCertificate -Server $Server -Thumbprint $Thumbprint -ErrorAction Stop
        Write-LogFile -Message "Loaded certificate properties from server $($Server)"
    }

    catch
    {
        Write-LogFile -Message "Failed to load certificate properties from server $($Server)" -ErrorInfo $_
    }

    If ($TLSCert -like "*")
    {
        Write-LogFile -Message "Building TLSCertname String"
        $TLSCertName = "<I>$($TLSCert.Issuer)<S>$($TLSCert.Subject)"
        Return $TLSCertName
        Write-LogFile -Message "Successfully Built TLSCertname String '$($TLSCertName)'"
    }
}

function Set-HybridCertificate
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $TLSCertName,
        [Parameter(Mandatory = $true)]
        [string]
        $SendConnector,
        [Parameter(Mandatory = $true)]
        [string[]]
        $Servers,
        [Parameter(Mandatory = $true)]
        [string]
        $DomainController
    )

    Foreach ($Server in $Servers)
    {
        try
        {
            Write-LogFile -Message "Loading properties of Receive Connector 'Default Frontend' from server $($Servers)"
            $RcvConnector = Get-ReceiveConnector -Server $Server | Where-Object Identity -Like "*Default Frontend*" -ErrorAction Stop
            Write-LogFile -Message "Loading properties of Receive Connector 'Default Frontend'"
        }
        catch
        {
            Write-LogFile -Message "Failed to load properties of Receive Connector 'Default Frontend' from server $($Server)" -ErrorInfo $_
            $SetHybridCertificateStatus = "Error"
        }

        Try
        {
            Write-LogFile -Message "Removing TLSCertificate value from Receive Connector $($RcvConnector.Identity)"
            Set-ReceiveConnector -Identity $RcvConnector.Identity -TlsCertificateName $null -DomainController $DomainController
            Write-LogFile -Message "Removed TLSCertificate value"
        }

        Catch
        {
            Write-LogFile -Message "Failed to remove TLSCertificate value from Receive Connector $($RcvConnector.Identity)" -ErrorInfo $_
            $SetHybridCertificateStatus = "Error"
        }

        Try
        {
            Write-LogFile -Message "Setting new value for TLSCertificate on Receive Connector $($RcvConnector.Identity)"
            Set-ReceiveConnector -Identity $RcvConnector.Identity -TlsCertificateName $TLSCertName -Confirm:$false -DomainController $DomainController
            Write-LogFile -Message "Set new value for TLSCertificate"
        }

        Catch
        {
            Write-LogFile -Message "Failed to set new value for TLSCertificate on Receive Connector $($RcvConnector.Identity)" -ErrorInfo $_
            $SetHybridCertificateStatus = "Error"
        }
    }

    try
    {
        Write-LogFile -Message "Removing TLSCertificate value from Send Connector $($SendConnector)"
        Set-SendConnector -Identity $SendConnector -TlsCertificateName $null -DomainController $DomainController -ErrorAction Stop
        Write-LogFile -Message "Removed TLSCertificate value"
    }
    
    catch
    {
        Write-LogFile -Message "Failed to remove TLSCertificate value from Send Connector $($SendConnector)" -ErrorInfo $_
        $SetHybridCertificateStatus = "Error"
    }

    try
    {
        Write-LogFile -Message "Setting new value for TLSCertificate on Send Connector $($SendConnector)"
        Set-SendConnector -Identity $SendConnector -TlsCertificateName $TLSCertName -Confirm:$false -DomainController $DomainController -ErrorAction Stop
        Write-LogFile -Message "Set new value for TLSCertificate"
    }
    
    catch
    {
        Write-LogFile -Message "Failed to set new value for TLSCertificate on Send Connector $($SendConnector)" -ErrorInfo $_
        $SetHybridCertificateStatus = "Error"
    }

    try
    {
        Write-LogFile -Message "Removing TLSCertificate value from Hybrid Configuration Object"
        Set-HybridConfiguration -TlsCertificateName $null -Confirm:$false -DomainController $DomainController -ErrorAction Stop
        Write-LogFile -Message "Removed TLSCertificate value"
    }
    
    catch
    {
        Write-LogFile -Message "Failed to remove TLSCertificate value from Hybrid Configuration Object" -ErrorInfo $_
        $SetHybridCertificateStatus = "Error"
    }

    try
    {
        Write-LogFile -Message "Setting new value for TLSCertificate on Hybrid Configuration Object"
        Set-HybridConfiguration -TlsCertificateName $TLSCertName -Confirm:$false -DomainController $DomainController -ErrorAction Stop
        Write-LogFile -Message "Set new value for TLSCertificate"
    }
    
    catch
    {
        Write-LogFile -Message "Failed to set new value for TLSCertificate on Hybrid Configuration Object" -ErrorInfo $_
        $SetHybridCertificateStatus = "Error"
    }

    Return $SetHybridCertificateStatus
}

#region XAML
#Form Start

Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration

[xml]$XAMLForm = @'

<Window x:Name="Replace_Hybrid_Certificate_Form" x:Class="WpfApp1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        mc:Ignorable="d"
        Title="Replace Hybrid Certificate" Height="385" Width="475" WindowStartupLocation="CenterScreen">
    <Grid x:Name="Form1" Margin="0,0,0,21">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="16*"/>
        </Grid.ColumnDefinitions>
        <Label x:Name="LabelStep1" Content="1" HorizontalAlignment="Left" Margin="14,11,0,0" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Grid.ColumnSpan="2"/>
        <Label x:Name="Label_Server_Selection" Content="Select Exchange Servers from the list below:" HorizontalAlignment="Left" Margin="3,14,0,0" VerticalAlignment="Top" Width="242" Grid.Column="1"/>
        <Button x:Name="Button_Retrieve_Certificates" Content="Retrieve Certificates" HorizontalAlignment="Left" Height="30" Margin="292,45,0,0" VerticalAlignment="Top" Width="120" Grid.Column="1"/>
        <Label x:Name="LabelStep2" Content="2" HorizontalAlignment="Left" Margin="268,43,0,0" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" RenderTransformOrigin="5.239,0.065" Grid.Column="1"/>
        <Label x:Name="Label_Certificate_Selection" Content="Select a certificate from the list below:" HorizontalAlignment="Left" Height="27" Margin="3,149,0,0" VerticalAlignment="Top" Width="238" Grid.Column="1"/>
        <Label x:Name="LabelStep3" Content="3" HorizontalAlignment="Left" Margin="12,145,0,0" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Grid.ColumnSpan="2"/>
        <ComboBox x:Name="Combobox_Certificate_Selection" HorizontalAlignment="Left" Height="25" Margin="8,179,0,0" VerticalAlignment="Top" Width="404" Grid.Column="1" SelectedIndex="0"/>
        <Label x:Name="LabelStep4" Content="4" HorizontalAlignment="Left" Margin="12,213,0,0" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Grid.ColumnSpan="2"/>
        <Label x:Name="Label_SendConnector_Selection" Content="Select a Send Connecor from the list below:" HorizontalAlignment="Left" Height="27" Margin="3,217,0,0" VerticalAlignment="Top" Width="238" Grid.Column="1"/>
        <ComboBox x:Name="Combobox_SendConnector_Selection" HorizontalAlignment="Left" Height="25" Margin="8,247,0,0" VerticalAlignment="Top" Width="404" Grid.Column="1" SelectedIndex="0"/>
        <Button x:Name="Button_Replace_Certificate" Content="Replace Certificate" HorizontalAlignment="Left" Height="30" Margin="150,291,0,0" VerticalAlignment="Top" Width="120" Grid.Column="1"/>
        <ListBox x:Name="Listbox_ExchangeServers" Grid.Column="1" HorizontalAlignment="Left" Height="98" Margin="8,44,0,0" VerticalAlignment="Top" Width="233"/>
        <Label x:Name="LabelStep5" Content="5" HorizontalAlignment="Left" Margin="126,289,0,0" VerticalAlignment="Top" FontSize="16" FontWeight="Bold" Grid.Column="1"/>
    </Grid>
</Window>

'@ -replace 'mc:Ignorable="d"', '' -replace "x:Name", 'Name' -replace '^<Win.*', '<Window' -replace 'x:Class="\S+"', ''

#Read XAMLForm
$form = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $XAMLForm))
$XAMLForm.SelectNodes("//*[@Name]") | Where-Object { Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name) }

# Icon definition
#
# Icon in Base64 format
[string]$IconB64 = @"
iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAACXBIWXMAAAsTAAALEwEAmpwYAAACO0lEQVR4nO2WvU8UQRiHL7uRS9ScEU1uFq0toNLG2FnYmCzBQv4FOws0u9DIkdhcAVHjGePuKQVWx1FBLIyFJORs1NsZxTPHiQw7i4r4kUCiBMJrdqN87ezeLCyh8N7kKXd/T2bmfWcSiUY1ilMqnjmjWvRju0WBS5l+V8u0y/fh0EJKMtiIbDjLsumAD4OtSiZ7krhTTYYLWDODgeFbRfq2hjslbvB2HsxdrCNAC0ICFgXVolkv3HReCIW75FlnbALtFoWWYbsiHL4XAieLDHYtkL5uHVI0klN0UjjVP2m35SpQj3PFqfgEFI1cUnQCUWjpIfEJII10RhVw2ZGAYV8WElBzNbjymK5zfqDKFThRtCOEs7XEw7lWIYFR8hO+Lq5A6cOSx7VhmytwZEh8BSST3eCefBQgMPL6R+gWnH0+Kx5usExg66EAgdr8b8hPLHhcuO3fguSjGMLDBCqffsGtZ188eGfgX0AyNwvH+iqAet7wV0vDG2M7zi2QTQea7tuQDgreBNJJNrLA/OIKTNSWPPqffuYKpLJTwm2LdCJ+CNVtbajeq3EFjt58H2F24LW0Xm4TElAEcAWaIwm4q4D9g0jpxh1Rw1H3DgU0wrkLMi8PIg3fdS+j473vbPen9Tg8MB2fwOaSDacQZbbvv0Dv5Kt9E5AMlvn7lhiPT8B0BgUF1qdb89VqCum4JNgF4Y/SA3l2WjbZdPCV6nyTTNv3LHclFJ0UkYaXA8JXFQ2PtWbeNoUKNOq/qj8zrPRuxNWyjAAAAABJRU5ErkJggg==
"@

# Convert icon to bitmap
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($IconB64)
$bitmap.EndInit()
$bitmap.Freeze()

# Add icon to form
$Form.Icon = $bitmap
#endregion

# Connect to Exchange
ConnectExchange | Out-Null

# Populate Exchange Server Listbox
$Script:ExchangeServers = Get-ExchangeServers
foreach ($Server in $Script:ExchangeServers)
{
    $Listbox_ExchangeServers.Items.Add($Server.Name) | Out-Null
}

$Listbox_ExchangeServers.SelectionMode = "Multiple"

# Populate list of Send Connectors
$SendConnectors = Get-SendConnector

if ($SendConnectors.Count -gt 0)
{
    foreach ($connector in $SendConnectors)
    {
        $ListItem = $connector.Name
        $Combobox_SendConnector_Selection.Items.Add($ListItem) | Out-Null
    }
}

Else
{
    $Combobox_SendConnector_Selection.Items.Add("No Send Connector found!") | Out-Null
    $Button_Replace_Certificate.IsEnabled = $false
}

$Button_Retrieve_Certificates.Add_Click(
    {
        $Script:SelectedExchangeServers = New-Object System.Collections.ArrayList
        
        foreach ($item in $Listbox_ExchangeServers.SelectedItems)
        {
            $Script:SelectedExchangeServers.Add($item) | Out-Null
        }

        $CertificateList = Get-CommonExchangeServerCertificates -ExchangeServers $Script:SelectedExchangeServers
        
        if ($CertificateList -gt 0)
        {
            Foreach ($Certificate in $CertificateList)
            {
                $Combobox_Certificate_Selection.Items.Add($Certificate) | Out-Null
            }
        }

        else
        {
            $Combobox_Certificate_Selection.Items.Add("No valid certificates found. See logfile for details...")
            $Button_Replace_Certificate.IsEnabled = $false  
        }
    }
)

$Button_Replace_Certificate.Add_Click(
    {
        $SelectedCertificateThumbprint = $Combobox_Certificate_Selection.SelectedItem.ToString().Split(",")[2]
        
        $SelectedConnector = $Combobox_SendConnector_Selection.SelectedItem
        
        $TLSCertificateString = Get-CertificateString -Thumbprint $SelectedCertificateThumbprint -Server $SelectedExchangeServers[0]

        $DomainController = ($ExchangeServers[0]).CurrentDomainControllers[0]

        $Result = Set-HybridCertificate -TLSCertName $TLSCertificateString -SendConnector $SelectedConnector -Servers $Script:SelectedExchangeServers -DomainController $DomainController

        if ($Result -eq "Error")
        {
            [System.Windows.Forms.MessageBox]::Show("An error occured. Please examine the log.", "Replacement failed", 0, [System.Windows.Forms.MessageBoxIcon]::Stop)
        }

        else
        {
            [System.Windows.Forms.MessageBox]::Show("Replacement successful. Please examine log for details", "Replacement success", 0, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
)

$Form.ShowDialog()
