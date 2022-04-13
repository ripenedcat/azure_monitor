param (
	[Parameter(mandatory = $false)]
	$WebHookData
)

# Setting ErrorActionPreference to stop script execution when error occurs
$ErrorActionPreference = 'Stop'
# Note: this is to force cast in case it's not of the desired type. Specifying this type inside before the param inside param () doesn't work because it still accepts other types and doesn't cast it to this type
$WebHookData = [PSCustomObject]$WebHookData

function Get-PSObjectPropVal {
    param (
        $Obj,
        [string]$Key,
        $Default = $null
    )
    $Prop = $Obj.PSObject.Properties[$Key]
    if ($Prop) {
        return $Prop.Value
    }
    return $Default
}

# If runbook was called from Webhook, WebhookData and its RequestBody will not be null
if (!$WebHookData -or [string]::IsNullOrWhiteSpace((Get-PSObjectPropVal -Obj $WebHookData -Key 'RequestBody'))) {
    throw 'Runbook was not started from Webhook (WebHookData or its RequestBody is empty)'
}

# Collect Input converted from JSON request body of Webhook
$RqtParams = ConvertFrom-Json -InputObject $WebHookData.RequestBody

if (!$RqtParams) {
    throw 'RequestBody of WebHookData is empty'
}

[string[]]$RequiredStrParams = @(
    'ResourceGroupName'
    'HostPoolName'
    "MaxUserSessionsThreshold"
)

[string[]]$InvalidParams = @($RequiredStrParams | Where-Object { [string]::IsNullOrWhiteSpace((Get-PSObjectPropVal -Obj $RqtParams -Key $_)) })
[string[]]$InvalidParams += @($RequiredParams | Where-Object { $null -eq (Get-PSObjectPropVal -Obj $RqtParams -Key $_) })

if ($InvalidParams) {
    throw "Invalid values for the following $($InvalidParams.Count) params: $($InvalidParams -join ', ')"
}


[string]$ResourceGroupName = $RqtParams.ResourceGroupName
[string]$HostPoolName = $RqtParams.HostPoolName
[double]$MaxUserSessionsThreshold = $RqtParams.MaxUserSessionsThreshold

#$ResourceGroupName = "TinaSub_WVD"
#$HostPoolName = "NewPool"

# Ensures you do not inherit an AzContext in your runbook
Disable-AzContextAutosave -Scope Process
    
#region azure auth, ctx
#Collect the credentials from Azure Automation Account Assets
$ConnectionAsset = Get-AutomationConnection -Name "AzureRunAsConnection"

# Azure auth
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

$MaxSessionLimit = (Get-AzWvdHostPool -Name $HostPoolName -ResourceGroupName $ResourceGroupName).MaxSessionLimit
# Personal Account has a default session limit as 1
if ($MaxSessionLimit -ge 9999){
	$MaxSessionLimit=1
}
$SessionHosts = @(Get-AzWvdSessionHost -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName)
$nSessionHosts = $SessionHosts.Count
[int]$nUserSessions = @(Get-AzWvdUserSession -ResourceGroupName $ResourceGroupName -HostPoolName $HostPoolName).Count

#[double]$MaxUserSessionsThreshold = 0.9
[int]$MaxUserSessionsThresholdCapacity = [math]::Floor($nSessionHosts * $MaxSessionLimit * $MaxUserSessionsThreshold)

$currentCapacityPercentage = [math]::Floor(100 * $nUserSessions/($nSessionHosts * $MaxSessionLimit) )
Write-Output "nUserSessions = $nUserSessions "
Write-Output "nSessionHosts = $nSessionHosts "
Write-Output "MaxSessionLimit = $MaxSessionLimit "
Write-Output "currentCapacityPercentage = $currentCapacityPercentage "

if ($currentCapacityPercentage -ge $MaxUserSessionsThreshold*100){
	Write-Output "currentCapacityPercentage=$currentCapacityPercentage is great than the MaxUserSessionsThreshold $MaxUserSessionsThreshold * 100%"
	$url = "https://prod-16.northcentralus.logic.azure.com:443/workflows/978a0277533a4d9eb77058e345f470f5/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=qvCtVpKWJPY-3zHpexQlVSa0GBjKHpHzst-FpylV_rA"
    $body = @{
        nUserSessions= $nUserSessions
        HostPoolName = $HostPoolName
        MaxUserSessionsThreshold = $MaxUserSessionsThreshold
		currentCapacityPercentage = $currentCapacityPercentage
	}
    Invoke-RestMethod -Method 'Post' -Uri $url -Body ($body|ConvertTo-Json) -ContentType "application/json"
	Write-Output "Triggered Logic App to send an alert email"
}
Write-Output "Finished"

