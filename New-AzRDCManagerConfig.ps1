[CmdletBinding()]
Param (
    [Switch]$UsePublicIP
)

$RDGOutputFile = "$env:TEMP\$Title.rdg"

$AllVMs = [Ordered]@{}
Get-AzSubscription | ForEach-Object {
    $sub = $_.Name
    Select-AzSubscription -SubscriptionObject $_
    $AllVMs.$sub = [Ordered]@{}
    Get-AzResourceGroup | ForEach-Object {
        $rg = $_.ResourceGroupName
        [System.Collections.ArrayList]$AllVMs.$sub.$rg = @()
        (Get-AzVM -ResourceGroupName $rg).Where{$_.StorageProfile.OsDisk.OsType -eq "Windows"} | ForEach-Object {
            $AllVMs.$sub.$rg.Add($_) | Out-Null
        }
    }
}


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

    $subNode = $newRDGFile.ImportNode($subobject.group,$true)
    [void]$MainNode.AppendChild($subNode)
    $subGroup = $newRDGFile.RDCMan.File.Group | Where-Object {$_.properties.name -eq $sub}

    ForEach ($rg in $AllVMs[$sub].Keys) {
        If ($AllVMs.$sub.$rg.Count -gt 0) {
            [XML]$rgobject = "
                <group>
                    <properties>
                        <expanded>true</expanded>
                        <name>$rg</name>
                    </properties>
                </group>
            "
            $rgNode = $newRDGFile.ImportNode($rgobject.group,$true)
            [void]$subGroup.AppendChild($rgNode)
            $rgGroup = $subGroup.group | Where-Object {$_.properties.name -eq $rg}

            Foreach ($VM in $AllVMs.$sub.$rg) {
                $vNic = $VM.NetworkProfile.NetworkInterfaces.Where{$_.Primary -eq $true -or $_.Primary -eq $null}.Id | Get-AzNetworkInterface
                If ($UsePublicIP) {
                    $vNic_public = (Get-AzResource -ResourceId $vNic.IpConfigurations.PublicIpAddress.Id)
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
                [XML]$vmObject = $vmInfo -f $IpAddress,$VM.Name,(New-Object PSObject -Property $VM.Tags | Out-String)
                $vmNode = $newRDGFile.ImportNode($vmObject.server,$true)
                [Void]$rgGroup.AppendChild($vmNode)

            }
        }
        Else {Write-Warning ("[{0}] The {1} RG does not contain any Windows VMs" -f $sub,$rg)}
    }
}

# Saves the RDG file to the specified output path
$newRDGFile.Save($RDGOutputFile)