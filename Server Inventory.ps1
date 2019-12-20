<#
.SYNOPSIS
    Add or update a virtual machine entry in server inventory
.DESCRIPTION
    Performs the following action on a Windows VM:
    Gathers information needed for Share Point entry
    If entry needs to be added, add VM info in Server Inventory
    If entry needs to be updated, udpate necessary info into Server Inventory
.PARAMETER [switch]Manual
    Choose whether manual entry or not to allow crawl of all vCenters 
.NOTES
  Version:        1.0
  Author:         <Joshua Dooling>
  Creation Date:  <08/15/2019>
  Purpose/Change: Automate Server Inventory entry
  
.EXAMPLE
    PS C:\> "Server Inventory.ps1"
    PS C:\> "Server Inventory.ps1" -Manual
#>


[CmdletBinding(DefaultParameterSetName="Manual")]
Param(
    [Parameter(Mandatory=$false,ParameterSetName="Manual")]
    [switch]$Manual
)

﻿#Requires -RunAsAdministrator

Import-Module ..\powershell-modules\Sharepoint\Sharepoint.psm1

Import-Module ..\powershell-modules\VMware\UCFvCenter.psm1

Import-Module ActiveDirectory

$EntriesAdded = @()
$EntriesUpdated = @()
$allVMs = @()

$Zones = @("226","225","227","228")

#region Tag Translation for Owner field
$Tag = @{
    "Example" = "Example-test"
}
#endregion

$vCenters = "vcenter1.site.company.com", "vcenter2.site.company.com"

$NetCreds = Get-Credential -Message "Enter NET Credentials (Add net\ before uid)"
$DevCreds = Get-Credential -Message "Enter DEV Credentials (Add dev\ before uid)"
$QACreds = Get-Credential -Message "Enter QA Credentials (Add qa\ before uid)"

$vCenterCreds = $NetCreds

$vCenters | Foreach-Object { Connect-VIServer -Server $_ -Credential $vCenterCreds | Out-Null }

$tenantID = 'Token'
$clientID = 'Key'
$clientSecret = 'pass'
$List = "List"
$Site = "SharePointsite-address"

Connect-Sharepoint -TenantID $tenantID -ClientID $clientID -ClientSecret $clientSecret

$allSPItems = (Get-SharepointList -List $List -Site $Site) | Where-Object {$_.Disposition -ne "Decommissioned"}

#region Auto look through vCenters option picked
if(-not $Manual){


    #region start crawl of vCenters
    foreach($vCenter in $vCenters){

        $allVMs = get-view -ViewType VirtualMachine -Server $vCenter | Select-Object Name, Guest, @{Name="ResourceID"; e={($_.ResourcePool.Type + "-" + $_.ResourcePool.Value)}} | Where-Object {$_.Guest.GuestFullName -like "*Windows*"}

        $VMCount = $allVMs.Count
        $i = 1

        #region Go through each VM in current vCenter and grab info
        foreach($VM in $allVMs){

            $Percent = [Math]::Round(($i / $VMCount) * 100)

            $Name = $VM.Name.ToString().Split('-')[0].Trim()

            Write-Progress -Activity "Quering VMs in $vCenter" -Status "$Percent% Complete" -PercentComplete $Percent -CurrentOperation "Currently working on $($Name)"

            $validatedObj = $allSPItems | Where-Object { $_.Server -match $Name }

            $loc = $vCenter

            $Parameters = @{}        
            $domain = ""

            #region Fill in Parameters
            if($VM.guest.hostname -match "\w+\.(.*)"){

                $domain = $Matches[1]

                #region AD Creds switch
                $ADCreds = switch -Wildcard ($domain){
                    
                    "*dev.company.com" {$DevCreds; continue}
                    "*qa.company.com" {$QACreds; continue}
                    "*company.com" {$NetCreds; continue}
                    
                }
            
                $ADCheck = Get-ADComputer -Server $domain -Identity $Name -Properties OperatingSystem, Description -Credential $ADCreds

                $Parameters["Model"] = "VMware Virtual Platform"
                $Parameters["Category"] = "Virtual"
                $Parameters["IPv4"] = ((Get-VM -Name "$($Name) - *").Guest.IPAddress | where-object {$_ -like "1*"}) -join ","
                $Parameters["OS"] = $AdCheck.OperatingSystem
                $Parameters["Domain"] = $domain
                $Parameters["Location"] = switch -Wildcard ($loc){

                    "vcenter1*" {"location1"; Break}
                    "vcenter2*" {"location2"; Break}
                    "vcenter3*" {"location3"; Break}
                }

                $Parameters["Building"] = switch($Parameters["Location"]){

                    "location1" {"Room1"; Break}
                    "location2" {"Room2"; Break}
                }


                #region Add or Update Sharepoint entry
                if(!$validatedObj){                
                
                    "$($VM.Name) is not present in Share Point, adding now..."

                    #Only change Role when Object is added into SP
                    if($ADCheck.Description){
                        $Parameters["Role"] = $AdCheck.Description
                    }elseif($null -eq $ADCheck.Description){
                        $Parameters["Role"] = "N/A"
                    }

                    if($Parameters["Location"] -eq "Room2"){
                        $Parameters["Security Zone"] = ([array]::IndexOf($Zones,($Parameters["IPv4"].ToString().Split('.')[1].Trim())))+1
                    }

                    #region Environment
                    $Environment = Get-ResourcePool -Server $vCenter | select Name, ID | Where-Object {$_.ID -like $vm.ResourcePoolId}

                    if($null -eq $Environment){

                        if($Environment -notlike "*Non*" -and $Environment -contains "*Production*"){
                            $Parameters["Environment"] = "Production"
                        }else{
                            $Environment = switch -Wildcard ($domain){
                                "*dev.company.com" {"Development"; continue}
                                "*qa.company.com" {"QA"; continue}
                                "*company.com" {"N/A"; continue}
                            }
                            $Parameters["Environment"] = $Environment
                        }

                    }elseif($Environment){

                        if($Environment -notlike "*Non*" -and $Environment -like "*Production*"){
                            $Parameters["Environment"] = "Production"
                        }else{
                            $Environment = switch -Wildcard ($domain){
                                "*dev.company.com" {"Development"; continue}
                                "*qa.company.com" {"QA"; continue}
                                "*company.com" {"N/A"; continue}
                            }
                            $Parameters["Environment"] = $Environment
                        }
                    }
                    #endregion

                    $departmentTag = Get-VM -Name $VM.Name | Get-TagAssignment -Category "Department"

                    if($departmentTag -ne $null){
                        $OwnerTag = $TagTranslate[$departmentTag.Tag.ToString().Split('/')[1].Trim()]

                        if($null -ne $OwnerTag){
                            $Parameters["Owner"] = $OwnerTag
                        }
                    }

                    $done = Add-SharepointListEntry -List $List -Site $Site -Server $Name @Parameters | Out-String
                    
                    if($done){
                        $doneObj = New-Object System.Object
                    
                        $doneObj | Add-Member -type NoteProperty -Name "VMName" -Value $Name
                        $doneObj | Add-Member -type NoteProperty -Name "Location" -Value $Parameters["Location"]

                        $EntriesAdded += $doneObj
                    
                    }             
                }elseif($validatedObj){
                
                    "$($VM.Name) is present in Share Point, updating now..."

                    $done = Update-SharepointListEntry -List $List -Site $Site -Server $Name @Parameters | Out-String
                    if($done){
                        $updatedObj = New-Object System.Object

                        $updatedObj | Add-Member -type NoteProperty -Name "VMName" -Value $Name
                        $updatedObj | Add-Member -type NoteProperty -Name "Location" -Value $Parameters["Location"]

                        $EntriesUpdated += $updatedObj
                    }
                }
                #endregion
            }
            #endregion

            $i++
        }
        #endregion
    }
    #endregion
}
#endregion

#region Manual section
elseif($Manual){

    do{
        $VMName = Read-Host("VM Name")

        #region Grab which vCenter VM is located in and make sure it exists
        foreach($vCenter in $vCenters){
            Write-Verbose "Trying $vCenter..."
            $VM = Get-VM -Name "$($VMName)*" -Server $vCenter
            if($VM){
                $loc = $vCenter
                break
            }
        }
        #endregion

        #region Start logic if VM exist
        if($VM){

            $VMInfo = get-view -ViewType VirtualMachine -Server $loc -Filter @{"Name"="$($VMName) -*"} | Select-Object Name, Guest, @{Name="ResourceID"; e={($_.ResourcePool.Type + "-" + $_.ResourcePool.Value)}} | Where-Object {$_.Guest.GuestFullName -like "*Windows*"}

            $Name = $VMInfo.Name.ToString().Split('-')[0].Trim()

            #Figure out how to call 'Get-SharepointList' function everytime a user enters a VM Name
               
            Connect-Sharepoint -TenantID $tenantID -ClientID $clientID -ClientSecret $clientSecret
            $validatedObj = (Get-SharepointList -List "List" -Site "SharePointSite") | Where-Object {$_.Server -match $Name}

            $Parameters = @{}        
            $domain = ""

            #region Fill in Parameters
            if($VMInfo.guest.hostname -match "\w+\.(.*)"){

                $domain = $Matches[1]

                #region AD Creds switch
                $ADCreds = switch -Wildcard ($domain){
                    
                    "*dev.company.com" {$DevCreds; continue}
                    "*qa.company.com" {$QACreds; continue}
                    "*company.com" {$NetCreds; continue}
                    
                }
            
                $ADCheck = Get-ADComputer -Server $domain -Identity $Name -Properties OperatingSystem, Description -Credential $ADCreds

                $Parameters["Model"] = "VMware Virtual Platform"
                $Parameters["Category"] = "Virtual"
                $Parameters["IPv4"] = ((Get-VM -Name "$($Name)*").Guest.IPAddress | where-object {$_ -like "1*"}) -join ","
                if($null -eq $ValidatedObj.OS){
                    if($null -ne $ADCheck.OperatingSystem){
                        $Parameters["OS"] = $AdCheck.OperatingSystem
                    }else{
                        $Parameters["OS"] = $VMInfo.Guest.GuestFullName
                    }
                }
                $Parameters["Domain"] = $domain
                $Parameters["Location"] = switch -Wildcard ($loc){

                    "vcenter1*" {"location1"; Break}
                    "vcenter2*" {"location2"; Break}
                    "vcenter3*" {"location3"; Break}
                }

                $Parameters["Building"] = switch($Parameters["Location"]){

                    "location1" {"Room1"; Break}
                    "location2" {"Room2"; Break}
                }


                #region Add or Update Sharepoint entry
                if(!$validatedObj){                
                
                    "$($Name) is not present in Share Point, adding now..."

                    #Only change Role if Object is being added into SP
                    if($ADCheck.Description){
                        $Parameters["Role"] = $AdCheck.Description
                    }elseif($null -eq $ADCheck.Description){
                        $Parameters["Role"] = "N/A"
                    }

                    if($Parameters["Location"] -eq "DSO"){
                        $Parameters["Security Zone"] = ([array]::IndexOf($Zones,($Parameters["IPv4"].ToString().Split('.')[1].Trim())))+1
                    }

                    #region Environment
                    $Environment = Get-ResourcePool -Server $loc | select Name, ID | Where-Object {$_.ID -like $VMInfo.ResourceId}

                    if($null -eq $Environment){

                        if($Environment -notlike "*Non*" -and $Environment -contains "*Production*"){
                            $Parameters["Environment"] = "Production"
                        }else{
                            $Environment = switch -Wildcard ($domain){
                                "*dev.company.com" {"Development"; continue}
                                "*qa.company.com" {"QA"; continue}
                                "*company.com" {"N/A"; continue}
                            }
                            $Parameters["Environment"] = $Environment
                        }

                    }elseif($Environment){

                        if($Environment -notlike "*Non*" -and $Environment -like "*Production*"){
                            $Parameters["Environment"] = "Production"
                        }else{
                            $Environment = switch -Wildcard ($domain){
                                "*dev.company.com" {"Development"; continue}
                                "*qa.company.com" {"QA"; continue}
                                "*company.com" {"N/A"; continue}
                            }
                            $Parameters["Environment"] = $Environment
                        }
                    }
                    #endregion

                    $departmentTag = Get-VM -Name $VMInfo.Name | Get-TagAssignment -Category "Department"

                    if($departmentTag -ne $null){
                        $OwnerTag = $TagTranslate[$departmentTag.Tag.ToString().Split('/')[1].Trim()]

                        if($null -ne $OwnerTag){
                            $Parameters["Owner"] = $OwnerTag
                        }
                    }

                    $done = Add-SharepointListEntry -List $List -Site $Site -Server $Name @Parameters | Out-String
               
                    if($done){
                        $doneObj = New-Object System.Object
                    
                        $doneObj | Add-Member -type NoteProperty -Name "VMName" -Value $Name
                        $doneObj | Add-Member -type NoteProperty -Name "Location" -Value $Parameters["Location"]
                        $doneObj | Add-Member -type NoteProperty -Name "Time" -Value (Get-Date -Format "MM/dd/yyyy HH:mm")

                        $EntriesAdded += $doneObj
                    
                    }elseif(!$done){
                        Write-Host "$($Name) was not added to Server Inventory"
                    }            
                }elseif($validatedObj){
                
                    "$($Name) is present in Share Point, updating now..."

                    $done = Update-SharepointListEntry -List $List -Site $Site -Server $Name @Parameters | Out-String

                    if($done){
                        $updatedObj = New-Object System.Object

                        $updatedObj | Add-Member -type NoteProperty -Name "VMName" -Value $Name
                        $updatedObj | Add-Member -type NoteProperty -Name "Location" -Value $Parameters["Location"]
                        $updatedObj | Add-Member -type NoteProperty -Name "Time" -Value (Get-Date -Format "MM/dd/yyyy HH:mm")

                        $EntriesUpdated += $updatedObj
                    }elseif(!$done){
                        Write-Host "$($Name) was not updated in Server Inventory"
                    }  
                }
                #endregion
            }
        }elseif(-not $vm){
            write-Host "$($VMName) not found"
        }
        #endregion

        $MoreVMs = Read-Host("Would you like to run another VM? ('Y', 'y', 'N', 'n')")

    }while($MoreVMs -like 'y')

    #region CSV Export
    if(($null -ne $EntriesAdded) -and ($null -ne $EntriesUpdated)){
    
        $option = Read-Host("Do you want to output CSVs? (Y/N)")

        if($option -like 'Y' -or $option -like 'y'){

            if($EntriesAdded){
                $EntriesAdded | Select-Object @{Name="VM Name"; E={$_.VMName}}, @{Name="vCenter"; E={$_.Location}}, @{Name="Time"; E={$_.Time}}, @{Name="Action"; E={"Added"}} | Export-Csv -Path ".\vm-spinfo SP-Entries.csv" -NoTypeInformation -Append
            }
            if($EntriesUpdated){
                $EntriesUpdated | Select-Object @{Name="VM Name"; E={$_.VMName}}, @{Name="vCenter"; E={$_.Location}}, @{Name="Time"; E={$_.Time}}, @{Name="Action"; E={"Updated"}} | Export-Csv -Path ".\vm-spinfo SP-Entries.csv" -NoTypeInformation -Append
            }

            Write-Host "`nCSV created with name 'vm-spinfo SP-Entries.csv'"
        }
    }
    #endregion
}


