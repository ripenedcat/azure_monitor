# ================Parameters================
# Specify the OU Path
$OUPathList = ,'OU=Delete,DC=Tinatest,DC=com'
$storageAccountName = "tinaprofile"
$storageAccountKey = "0XNhVvorJH9TtHitGxQbonUJSZHVWGzwBqBCFpC5+CjsvUZ0HqfOsTcOoJ205nFR/5UBBCbfFGRkUR44jvh6cQ=="
$fileShareName = "profiles"
# ================End Of Parameters================

#stop script on any errors
$errorActionPreference = "Stop"
Disable-AzContextAutosave -Scope Process | Out-Null

#initialize some parameters
$finalProcessedInfomation = ""
$disabledUsersList = @()
$fileShareFolderNames = @()

# iterate all OU provided
foreach($OUPath in $OUPathList){
	if([String]::IsNullOrEmpty($OUPath)){
		continue
	}
	Write-Output "Checking disabled users under $OUPath"
	# List all disabled users from OU path
	$disabledUsers = Get-ADUser -Filter {Enabled -eq "false"} -SearchBase $OUPath | Select Name,Enabled,SamAccountName,SID,UserPrincipalName
	if ($disabledUsers.count -gt 0){
		# found disabled user, take a record
		$disabledUsersList+=$disabledUsers
		foreach($disabledUser in $disabledUsers){
			Write-Output "Disabled user found: $($disabledUser.Name) --- $($disabledUser.UserPrincipalName)"
		}
		Write-Output ""		
	}else{
		Write-Output "No disabled User found in this OU."
		Write-Output ""	
	}
}
# there is no disabled users in all OU.
if ($disabledUsersList.count -eq 0){
	Write-Output "Checked all OU. There is no disabled user found. Aborting."; 
	exit
}
Write-Output "Totally found $($disabledUsersList.count) disabled user(s)."
Write-Output ""	
# foreach($disabledUser in $disabledUsersList){
# 	Write-Output "$($disabledUser.Name) --- $($disabledUser.UserPrincipalName)"
# }

# concat fileshare folder name per disabled user's sid and account name
foreach($disabledUser in $disabledUsersList){
	#concat azure file share name
	$fileShareFolderNames += $disabledUser.SamAccountName + "_"+  @($disabledUser.SID).Value
}

Write-Output "Starting to delete fileshare per disabled user SID & Account Name"
Write-Output ""	
$storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccountKey
foreach($fileShareFolderName in $fileShareFolderNames){
	Write-Output "Deleting $fileShareFolderName"
	#Get-AzStorageFile -ShareName $fileShareName -Path $fileShareFolderName -Context $storageContext | Get-AzStorageFile | Remove-AzStorageFile
	try{
		#delete all files under the fileshare folder
		Get-AzStorageFile -ShareName $fileShareName -Path $fileShareFolderName -Context $storageContext | Get-AzStorageFile | Remove-AzStorageFile
		# delete the file share filder itself.
		Remove-AzStorageDirectory -ShareName $fileShareName -Path $fileShareFolderName -Context $storageContext 
		$finalProcessedInfomation += $fileShareFolderName +"`n"
	}catch{
		Write-Output "Delete failure on $fileShareFolderName"
		Write-Output ""	
		continue
	}
	Write-Output "Successfully deleted $fileShareFolderName"
	Write-Output ""	
}
Write-Output ""	
Write-Output "==============In Summarize=============="
Write-Output "Within the provided OU(s), there are $($disabledUsersList.count) disabled user(s) found."
Write-Output ""	
if([String]::IsNullOrEmpty($finalProcessedInfomation)){
	Write-Output "No Fileshare was successfully deleted."
	$finalProcessedInfomation = "No Fileshare was successfully deleted."
}else{
	Write-Output "Fileshare successfully deleted:"
	Write-Output "$finalProcessedInfomation"
}
$url = ""
$body = @{
	finalProcessedInfomation= $finalProcessedInfomation
}
Invoke-RestMethod -Method 'Post' -Uri $url -Body ($body|ConvertTo-Json) -ContentType "application/json"
Write-Output "Triggered Logic App to send an email"


Write-Output ""
Write-Output "Script finished"
