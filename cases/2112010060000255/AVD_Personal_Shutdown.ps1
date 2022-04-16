
#==============Parameters=================
#your environment information here.
$ResourceGroupName = "TinaSub_WVD"
$HostPoolName = "Personal"
#Put exclusion VM Name (without domain) in the list. This is for some VMs which are not in a AAD Group, but still want to be excluded to be turned off. 
$VMExclusionList = ,""
#Put AAD Group name here. all VMs in this AAD Group will not be turned off.
$AADGroupExclusionNameList= ,""
#=========================================


$errorActionPreference = "Stop"
# Initialize an empty variable for transfer AAD group name to id 
$AADGroupExclusionIDList= @()
# Initialize an empty array to record VM turned off by this script.
$DeallocatedVMList= @()
#==Section of Azure Login==
$ConnectionAsset = Get-AutomationConnection -Name "AzureRunAsConnection"
$AzContext = $null
try {
    $AzAuth = Connect-AzAccount -ApplicationId $ConnectionAsset.ApplicationId -CertificateThumbprint $ConnectionAsset.CertificateThumbprint -TenantId $ConnectionAsset.TenantId -SubscriptionId $ConnectionAsset.SubscriptionId  -ServicePrincipal
    if (!$AzAuth -or !$AzAuth.Context) {
        throw $AzAuth
    }
    $AzContext = $AzAuth.Context
}
catch {
    throw [System.Exception]::new('Failed to authenticate Azure with application ID, tenant ID, subscription ID', $PSItem.Exception)
}
Write-Output "Successfully authenticated with Azure using service principal: $($AzContext | Format-List -Force | Out-String)"
#===============================
#==Section of Azure AD Login==
try {
    $AzureADAuth = Connect-AzureAD -ApplicationId $ConnectionAsset.ApplicationId -CertificateThumbprint $ConnectionAsset.CertificateThumbprint -TenantId $ConnectionAsset.TenantId
}
catch {
    throw [System.Exception]::new('Failed to authenticate Azure AD with application ID, tenant ID, subscription ID', $PSItem.Exception)
}
Write-Output "Successfully connected to Azure AD"
#==================================


#==Add AAD Group Exception VMs to $VMExclusionList
foreach($AADGroup in $AADGroupExclusionNameList){
	if([String]::IsNullOrEmpty($AADGroup)){
		continue
	}
	$filter = "DisplayName eq '"+ $AADGroup +"'"
	$AADGroupObjectId = (Get-AzureADGroup -Filter $filter).ObjectId
	$AADGroupExclusionIDList+=$AADGroupObjectId
}
foreach($AADGroupExclusionID in $AADGroupExclusionIDList){
	if($AADGroupExclusionID -eq ""){
		continue
	}
	$VMNames= (Get-AzureADGroupMember -ObjectId $AADGroupExclusionID).DisplayName
	foreach($VMName1 in $VMNames){
		$VMExclusionList += $VMName1
	}
}
Write-Output("Adding Single VMs exception and AAD group exception together, the Exception VMs are $VMExclusionList")
#============================


Write-Output("Start Scanning Hosts in verbose")
$SessionHosts = Get-AzWvdSessionHost -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName
#Iterate all VMs in a session Hosts
foreach($SessionHost in $SessionHosts){
	
	$hostName = ($SessionHost.Name -split "/")[1]
	#get current user session
	$userSession = Get-AzWvdUserSession -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName -SessionHostName $hostName
	#get current user session count
	$sessionCount = ($userSession | measure).count
	# shutdownFlag by default is true. true means VM will be turned off
	$shutdownFlag = $true
	# initialize sessionState as N/A when there is no current user Session
	$sessionState = "N/A"
	if ($sessionCount -gt 0){
		#get real session state. if there is a session with Active or Disconnected state, set $shutdownFlag to false and VM will not be turned off
		$sessionState = $userSession.SessionState
		if (($sessionState -eq "Active") -or ($sessionState -eq "Disconnected")){			
			$shutdownFlag = $false
		}
	}
	Write-Output "The session host $hostName has $sessionCount session, session state is $sessionState, will be turned off with $shutdownFlag."
	# check if this VM will be shutdown
	if ($shutdownFlag){
		$VMName = ($hostName -split "\.")[0]
		#if the VM is in VMExclusionList, skip turning off. or will start deallocating
		if($VMExclusionList -contains $VMName){
			Write-Output "$VMName is in VMExclusionList, will skip deallocate"
			continue
		}else{
			Write-Output "$VMName is not in VMExclusionList, will start deallocation"
		}
		try{
			# check current VM Status, if already deallocated, no need to deallocate again
			$VMStatus = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName -Status			
			if ($VMStatus.Statuses.Code -contains ("PowerState/deallocated")){
				Write-Output "$VMName is already deallocated, skip"
			}else{
				$a = Stop-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName -Force -NoWait
				$DeallocatedVMList += $VMName
				Write-Output "VM $VMName is being deallocated."
			}
			
		}catch{
			Write-Output "Encounterred an error when deallocating the VM $VMName."
		}
		
	}

}
# Output Summary
Write-Output "Script Finished."
if ( $DeallocatedVMList.count -eq 0){
	Write-Output "In summarize, no VM is turned off by this script."
}else{
	Write-Output "In summarize, VMs turned off by this script are $DeallocatedVMList"
}

