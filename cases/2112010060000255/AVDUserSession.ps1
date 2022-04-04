#your environment information here.
$ResourceGroupName = "TinaSub_WVD"
$HostPoolName = "Personal"
#Put exclusion VM Name (without domain) in the list 
$VMExclusionList = ,"Per-0"
#$ADGroupExclusionIDList= "","2ca0dd39-8195-4531-9e7a-95a5e21bb403"
$AADGroupExclusionNameList= ,"AAD add gp","Computer_aadgp"
$AADGroupExclusionIDList= @()


#==Azure Login==
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

#==Azure AD Login==
try {
    $AzureADAuth = Connect-AzureAD -ApplicationId $ConnectionAsset.ApplicationId -CertificateThumbprint $ConnectionAsset.CertificateThumbprint -TenantId $ConnectionAsset.TenantId
}
catch {
    throw [System.Exception]::new('Failed to authenticate Azure AD with application ID, tenant ID, subscription ID', $PSItem.Exception)
}
Write-Output "Successfully connected to Azure AD"

#==Adding AAD Group Exception VMs to $VMExclusionList
foreach($AADGroup in $AADGroupExclusionNameList){
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

Write-Output("Start Scanning Hosts")
$SessionHosts = Get-AzWvdSessionHost -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName

foreach($SessionHost in $SessionHosts){
	$hostName = ($SessionHost.Name -split "/")[1]
	
	$userSession = Get-AzWvdUserSession -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName -SessionHostName $hostName
	$shutdownFlag = $true
	$sessionCount = ($userSession | measure).count
	$sessionState = "N/A"
	if ($sessionCount -gt 0){
		$sessionState = $userSession.SessionState
		if (($sessionState -eq "Active") -or ($sessionState -eq "Disconnected")){
			#if (($sessionState -eq "Active")){
			$shutdownFlag = $false
		}
	}
	Write-Output "The session host $hostName has $sessionCount session, its state is $sessionState, will be turned off with $shutdownFlag."
	if ($shutdownFlag){
		$VMName = ($hostName -split "\.")[0]
		if($VMExclusionList -contains $VMName){
			Write-Output "$VMName is in VMExclusionList, will skip deallocate"
			continue
		}else{
			Write-Output "$VMName is not in VMExclusionList, will start deallocation"
		}
		try{
			$VMStatus = Get-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName -Status
			#$b.Statuses
			if ($VMStatus.Statuses.Code -contains ("PowerState/deallocated")){
				Write-Output "$VMName is already deallocated, skip"
			}else{
				$a = Stop-AzVM -ResourceGroupName $ResourceGroupName -Name $VMName -Force -NoWait
				Write-Output "VM $VMName is being deallocated."
			}
			
		}catch{
			Write-Output "Encounterred an error when deallocating the VM $VMName."
		}
		
	}

}

Write-Output "Finished"

