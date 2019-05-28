[CmdletBinding()]
Param (
    [Parameter(Mandatory=$true)]
    [String]$SubscriptionName = "ALL",
    [Parameter(Mandatory=$true)]
    [String]$Title,
    $storageAccountName,
    $storageAccountContainerName,
    $storageAccountResourceGroupName,
    $storageAccountSubscription,
    [Switch]$UsePublicIP
)

try {
    If ((Get-Command Get-AutomationConnection -ErrorAction SilentlyContinue)) {
        $AzureAutomation = $true
        try {
            Write-Verbose "Found Azure Automation commands, checking for Azure RunAs Connection..."
            # Attempts to use the Azure Run As Connection for automation
            $svcPrncpl = Get-AutomationConnection -Name "AzureRunAsConnection"
            $tenantId = $svcPrncpl.tenantId
            $appId = $svcPrncpl.ApplicationId
            $crtThmprnt = $svcPrncpl.CertificateThumbprint
            Add-AzureRMureRmAccount -ServicePrincipal -TenantId $tenantId -ApplicationId $appId -CertificateThumbprint $crtThmprnt -EnvironmentName AzureUsGovernment | Out-Null
        }
        catch {Write-Error -Exception "Azure RunAs Connection Failure" -Message "Unable to use Azure RunAs Connection" -Category "OperationStopped" -ErrorAction Stop}
    }
    Else {Write-Verbose ("Azure Automation commands missing, skipping Azure RunAs Connection...")}
    
    If ($Title.Contains(" ")) {$Title = $Title.Replace(" ","_")}
    $RDGOutputFile = "$env:TEMP\$Title.rdg"

    Write-Output ("--------------------- Start VM Collection Process ---------------------`n`r")
    $AllVMs = [Ordered]@{}
    If ($SubscriptionName -eq "ALL") {
        Write-Verbose "Collecting Virtual Machines from All Azure Subscriptions"
        Get-AzureRMSubscription | ForEach-Object {
            $sub = $_.Name
            Write-Output ("[{0}] Connecting to Azure Subscription" -f $sub)
            $AzureContext = Select-AzureRMSubscription -SubscriptionObject $_
            If ($AzureContext) {Write-Output ("Connected to: {0}" -f $AzureContext.Name)}
            Else {
                Write-Output ("[{0}] - Azure Context not found!" -f $sub)
                Write-Error -Exception "Invalid Azure Context" -Message ("Unable to create an Azure Context under the {0} subscription" -f $sub) -Category "OperationStopped" -ErrorAction Stop
            }
            $AllVMs.$sub = [Ordered]@{}
            Get-AzureRMResourceGroup | ForEach-Object {
                $rg = $_.ResourceGroupName
                Write-Output ("[{0}] - Working on {1} Resource Group" -f $sub,$rg)
                [System.Collections.ArrayList]$AllVMs.$sub.$rg = @()
                (Get-AzureRMVM -ResourceGroupName $rg).Where{$_.StorageProfile.OsDisk.OsType -eq "Windows"} | ForEach-Object {
                    $AllVMs.$sub.$rg.Add($_) | Out-Null
                }
            }
        }
    }
    Else {
        Write-Verbose ("Collecting Virtual Machines from {0} Azure Subscription" -f $SubscriptionName)
        Write-Output ("[{0}] Connecting to Azure Subscription" -f $sub)
        $sub = $SubscriptionName
        $AzureContext = Select-AzureRMSubscription -Subscription $sub
        If ($AzureContext) {Write-Output ("Connected to: {0}" -f $AzureContext.Name)}
        Else {
            Write-Output ("[{0}] - Azure Context not found!" -f $sub)
            Write-Error -Exception "Invalid Azure Context" -Message ("Unable to create an Azure Context under the {0} subscription" -f $sub) -Category "OperationStopped" -ErrorAction Stop
        }
        $AllVMs.$sub = [Ordered]@{}
        Get-AzureRMResourceGroup | ForEach-Object {
            $rg = $_.ResourceGroupName
            Write-Output ("[{0}] - Working on {1} Resource Group" -f $sub,$rg)
            [System.Collections.ArrayList]$AllVMs.$sub.$rg = @()
            (Get-AzureRMVM -ResourceGroupName $rg).Where{$_.StorageProfile.OsDisk.OsType -eq "Windows"} | ForEach-Object {
                $AllVMs.$sub.$rg.Add($_) | Out-Null
            }
        }
    }
    Write-Output ("--------------------- End VM Collection Process ({0}) ---------------------`n`r" -f $stopwatch.Elapsed)
    $stopwatch.Restart()
    Write-Output "---------------------- Start RDG File Creation Process ----------------------"
    Write-Output ("Building RDG file using XML...")
    # Build the New RDG file as XML
    # NOTE: THIS XML IS SPECIFIC TO RDCMAN VERSION 2.7
    $newRDGFile = @"
<?xml version="1.0" encoding="utf-8"?>
<RDCMan programVersion="2.7" schemaVersion="3">
    <file>
        <credentialsProfiles />
        <properties>
            <expanded>True</expanded>
            <name>$Title</name>
        </properties>
        <remoteDesktop inherit="None">
        <sameSizeAsClientArea>True</sameSizeAsClientArea>
        <fullScreen>False</fullScreen>
        <colorDepth>32</colorDepth>
        </remoteDesktop>
        <localResources inherit="None">
        <audioRedirection>NoSound</audioRedirection>
        <audioRedirectionQuality>Dynamic</audioRedirectionQuality>
        <audioCaptureRedirection>DoNotRecord</audioCaptureRedirection>
        <keyboardHook>FullScreenClient</keyboardHook>
        <redirectClipboard>True</redirectClipboard>
        <redirectDrives>False</redirectDrives>
        <redirectDrivesList />
        <redirectPrinters>False</redirectPrinters>
        <redirectPorts>False</redirectPorts>
        <redirectSmartCards>True</redirectSmartCards>
        <redirectPnpDevices>False</redirectPnpDevices>
        </localResources>
    </file>
    <connected />
    <favorites />
    <recentlyUsed />
</RDCMan>
"@ -as [XML]

    $MainNode = $newRDGFile.RDCMan.file

    Foreach ($sub in $AllVMs.Keys) {
        [XML]$subobject = "
            <group>
                <properties>
                    <expanded>true</expanded>
                    <name>$sub</name>
                </properties>
            </group>
        "

        Write-Output ("[{0}] Connecting to Azure Subscription" -f $sub)
        $AzureContext = Select-AzureRMSubscription -Subscription $sub
        If ($AzureContext) {Write-Output ("Connected to: {0}" -f $AzureContext.Name)}
        Else {
            Write-Output ("[{0}] - Azure Context not found!" -f $sub)
            Write-Error -Exception "Invalid Azure Context" -Message ("Unable to create an Azure Context under the {0} subscription" -f $sub) -Category "OperationStopped" -ErrorAction Stop
        }

        $subNode = $newRDGFile.ImportNode($subobject.group,$true)
        [void]$MainNode.AppendChild($subNode)
        $subGroup = $newRDGFile.RDCMan.File.Group | Where-Object {$_.properties.name -eq $sub}

        ForEach ($rg in $AllVMs[$sub].Keys) {
            If ($AllVMs.$sub.$rg.Count -gt 0) {
                [XML]$rgobject = "
                    <group>
                        <properties>
                            <expanded>false</expanded>
                            <name>$rg</name>
                        </properties>
                    </group>
                "
                Write-Output ("[{0}] - Working on {1} Resource Group ({2} VMs)" -f $sub,$rg,$AllVMs.$sub.$rg.Count)
                $rgNode = $newRDGFile.ImportNode($rgobject.group,$true)
                [void]$subGroup.AppendChild($rgNode)
                $rgGroup = $subGroup.group | Where-Object {$_.properties.name -eq $rg}

                Foreach ($VM in $AllVMs.$sub.$rg) {
                    $vNic = $VM.NetworkProfile.NetworkInterfaces.Where{$_.Primary -eq $true -or $_.Primary -eq $null}.Id | Get-AzureRMNetworkInterface
                    If ($UsePublicIP) {
                        $vNic_public = (Get-AzureRMResource -ResourceId $vNic.IpConfigurations.PublicIpAddress.Id)
                        $IpAddress = $vNic_public.Properties.IPAddress
                    }
                    Else {$IpAddress = $vNic.IpConfigurations.PrivateIpAddress}

                    $vmInfo = "
                        <server>
                            <name>{0}</name>
                            <displayName>{1}</displayName>
                            <comment>{2}</comment>
                        </server>
                    "
                    [System.Collections.ArrayList]$Tagdata = @()
                    $VM.Tags.Keys | Sort-Object | ForEach-Object {
                        [Void]$Tagdata.Add(("{0}: {1}&#13;`r" -f $_,$VM.Tags[$_]))
                    }
                    [XML]$vmObject = $vmInfo -f $IpAddress,$VM.Name,($Tagdata | Out-String)
                    $vmNode = $newRDGFile.ImportNode($vmObject.server,$true)
                    [Void]$rgGroup.AppendChild($vmNode)

                }
            }
            Else {Write-Warning ("[{0}] The {1} RG does not contain any Windows VMs" -f $sub,$rg)}
        }
    }

    # Saves the RDG file to the specified output path
    $newRDGFile.Save($RDGOutputFile)
    Write-Output ("--------------------- End RDG File Creation Process ({0}) ---------------------`n`r" -f $stopwatch.Elapsed)

    If ([String]::IsNullOrEmpty($storageAccountName) -or [String]::IsNullOrEmpty($storageAccountContainerName) -or [String]::IsNullOrEmpty($storageAccountSubscription) -or [String]::IsNullOrEmpty($storageAccountResourceGroupName)) {
        If ($AzureAutomation) {Write-Warning ("Azure Automation detected, but missing Azure Storage parameters.  RDG file was written to: {0}" -f $RDGOutputFile)}
        Else {Write-Output ("RDG file was created and saved to: {0}" -f $RDGOutputFile)}
    }
    Else {
        $stgAccount = Get-AzureRMStorageAccount -ResourceGroupName $storageAccountResourceGroupName -Name $storageAccountName
        Set-AzureStorageBlobContent -File $RDGOutputFile -Container $storageAccountContainerName -Blob "$Title.rdg" -Context $stgAccount.Context
        Write-Output ("RDG File was created and saved to: {0}\{1}\{2}\{3}\{4}\{5}.rdg" -f $storageAccountSubscription,$storageAccountResourceGroupName,$storageAccountName,$storageAccountContainerName,$Title)
    }
}
catch {$PSCmdlet.ThrowTerminatingError($PSItem)}