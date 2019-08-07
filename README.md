# Az.RDPConnectionBuilder

## Overview
This repository contains a single PowerShell script which can be used standalone or part of an Azure Automation Runbook.  The script will attempt to find all Windows Virtual Machines in every Azure Subscription you have permissions to see.  After collecting all the Windows Virtual Machines, it will create a custom RDG file (used with RDC Manager 2.7) allowing easy RDP access to all the Virtual Machines in your subscriptions.  The script will also parse the tags from the Virtual Machines and add them as comments on the servers in the RDG file.

## Script
``New-AzRDCManagerConfig.ps1``

### Examples
#### ALL Azure Subscriptions to local file - Private IP Addresses
> ``New-AzRDCManagerConfig.ps1 -SubscriptionName ALL -Title "AzureVMs"``
> - Creates a RDG file in the local %temp% folder called **AzureVMs.rdg**
> - Add the **switch** ``-UsePublicIP`` to have the config file built using the public IP instead of the private IP.

#### Specific Subscription to Storage Account (Blob container) - Private IP Addresses
> ``New-AzRDCManagerConfig.ps1 -SubscriptionName PROD -Title AzureVMs -storageAccountName carbonstgacct42 -storageAccountContainerName RDCManager -storageAccountResourceGroupName lab-carbon-rg -storageAccountSubscription PROD``
> - Creates a RDG file in the *RDCManager* container of the *carbonstgacct42* Storage Account called **AzureVMs.rdg**
> - Add the **switch** ``-UsePublicIP`` to have the config file built using the public IP instead of the private IP.

## Requirements
- RDC Manager v2.7 ([Download](https://www.microsoft.com/en-us/download/details.aspx?id=44989))
- **Supports Azure Automation**
