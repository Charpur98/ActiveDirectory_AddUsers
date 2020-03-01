Import-Module ActiveDirectory

#File Paths
$UserDataBase 	= "C:\Useras\Charpur\Desktop\PowerShell\usersDB.csv"
$User_DB_BackUp = "C:\Users\Charpur\Desktop\PowerShell\usersDB_BackUp.csv"
$UserImport 	= "C:\Users\Charpur\Desktop\PowerShell\usersDB.csv"
$PhoneNumsFile 	= "C:\Users\Charpur\Desktop\PowerShell\phoneNums.csv"
$PS_Log 		= "C:\Users\Charpur\Desktop\PowerShell\PS_Log.txt"

$SenderEmail = Get-Content "C:\Users\Charpur\Desktop\EmailRequests\Data\SenderEmail.txt"

$Date = Get-Date -Format "dd/MM/yy HH:mm"



function startfunc(){

	SetUpDB
	LoadUsers
	write "Finished."
}


function LoadUsers(){
	$UserFileExists = Test-Path $UserImport
	write "Loading users..."
	If ($UserFileExists -eq $True) {
		$ADUsers = Import-csv $UserImport
	}
	Else { 
		write "Load User Fail"
		Out-File $PS_Log -Append -InputObject "Error: Loading Users  |  $Date "
		exit #End Script if fail to load .csv file
	} 

	if([string]::IsNullOrEmpty($ADUsers)){
		write "Load User Fail - File is empty."
		Out-File $PS_Log -Append -InputObject "Warning: Loading Users - No users to add | $Date "
		exit
	}
 

	$UserIndex = 0
	foreach ($User in $ADUsers){
		#Stored Var	| #Column heading from .csv
		$ADProgress	= $User.AD_progress
		$WinAProgress = $User.WinA_Progress
			if($ADProgress -eq "Created"){ $UserIndex++; continue} #Exit loop if account is created
			if($ADProgress -eq "FAILED"){ $UserIndex++; continue} #Exit loop if account is created

		$Manager 	= $User.manager
		if((CheckAccountExists $Manager) -eq "false"){	#Check if manager exists in the AD
			Out-File $PS_Log -Append -InputObject "Warning: $Manager doesnt exist | $Date "
			$Manager = $null
		} else{
			$ManagerEMail 	= Get-ADUser -Identity $Manager | select "UserPrincipalName"; $ManagerEMail = $ManagerEMail.UserPrincipalName
		}

		$Firstname  = CheckValidString $User.firstname
		$Middlename = $User.middlename
		$Lastname   = CheckValidString $User.lastname
		$Name 		= "$Firstname $Lastname"
		$Username	= CreateUserName $Firstname $Lastname
		$Email 		= CreateUserEmail $Username
		$Department = $User.department
		$JobTitle 		= $User.jobtitle
			$JobTitle = $JobTitle.Replace(",","-")
		$Office			= $User.office

		$AccountExists		= CheckAccountExists $Username
		$TempPassword 		= RandString
		$UserPhoneNumber 	= AssignPhoneNum

		if($AccountExists -eq "false"){
			#Create the AD Account
			$ADProgress = AddActiveDir	
			if ($ADProgress -eq "Created"){
				$User.UserName 		= $Username
				$User.Email 		= $Email
				$User.PhoneNumber 	= $UserPhoneNumber
				$User.TempPassword 	= $TempPassword
				$User.Status 		= $Status = "Active"
				$User.AD_Progress 	= "Created"
				$User.WinA_Progress	= "Requested"

				$User
				SendEmail_UserCreated
				$AccountCreated = $true
				
			}
			
		}
		else{
			Out-File $PS_Log -Append -InputObject "Warning: $Username exists, check if the user has been already created or there may be a username conflict.  | $Date"
			SendEmail_Error "Failed to create account for $firstname $lastname, $Username already exists."
			$User.AD_Progress = "FAILED"
		}
		$UserIndex++
	}
	write "-----------------------------------"
	ExportToDB
		if($AccountCreated -eq $true){
		}
		else { write "No accounts created."}
}


function CheckValidString([string] $CheckString){

	$CheckString = $CheckString.Replace("-"," ")

	if($CheckString -match '^[a-z ]+$'){

		return $CheckString
	}
	else {
		Out-File $PS_Log -Append -InputObject "Error: Invalid Name: $CheckString | $Date "
		$firstname = $User.firstname; $lastname = $user.lastname
		SendEmail_Error "Failed to create account for $firstname $lastname, the name contains invalid characters."
		write "Invalid Name: $CheckString!"

		$User.AD_Progress = "FAILED"
		$UserIndex++
		continue
	}
}


function CreateUserName([string]$firstName, [string]$lastName){

	$Firstchar 	= $firstName.Substring(0,1)
	$Username 	= "$Firstchar$lastName"

	$Username = $Username.Replace(" ","")
	$Username = $Username.Replace("-","")

	
	return $Username.ToLower()

}


function CreateUserEmail([string]$Username){

	$UserEmail = $Username+"@domain.com"

	return $UserEmail

}


#Check that the username doesnt already exists in AD
function CheckAccountExists([string]$CheckUsername){

	Try {
		Get-ADUser -Identity $CheckUsername;
		$AccountExists = "true"
		return $AccountExists
	}
	Catch {
		 $AccountExists = "false"
		return $AccountExists
	}

}


function RandString(){

	$Chars 		= "abcdefghijklmnopqrstuvwxyz"
	$UC_chars 	= "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	$Nums 		= "1234567890"

	for($i = 0; $i -lt 9; $i++){
		$charRand = Get-Random -Maximum 3
		if($charRand -eq 1){
			$randchar = Get-Random -Maximum $UC_chars.Length
			$UserPassword += $UC_chars[$randchar]
			$char1 = "true"
		}
		if($charRand -eq 2){
			$randchar = Get-Random -Maximum $Nums.Length
			$UserPassword += $Nums[$randchar]
			$char2 = "true"		
		}
		else{
			$randchar = Get-Random -Maximum $Chars.Length
			$UserPassword += $Chars[$randchar]
			$char3 = "true"
		}
	}	
	#Ensure the password meets the complexity requirements
	if ($char1 -eq "true" -and $char2 -eq "true" -and $char3 -eq "true"){
		return $UserPassword
	}
	else {RandString}
}


function AddActiveDir(){

	$Password = ConvertTo-SecureString -String $TempPassword -AsPlainText -Force

	Try{
		New-ADUser	`
			-Name 				$Name `
			-GivenName 			$Firstname `
			-Othername 			$Middlename `
			-Surname 			$Lastname `
			-DisplayName		$Name `
			-SamAccountName 	$Username `
			-UserPrincipalName 	$Email `
			-EmailAddress 		$Email `
			-AccountPassword 	$Password `
			-Department 		$Department `
			-Office				$Office `
			-Manager 			$Manager `
			-Description  		$JobTitle `
			-Title  			$JobTitle `
			-OfficePhone		$UserPhoneNumber `
			-Company			"CompanyName" `
			-OtherAttributes @{'EmployeeType'="$EmployeeType"} `
			-Enabled			$true `
			-Path 				$ADLoc

			AddAddress $Username
			#Once account is created, add groups
			SetADGroups

		Out-File $PS_Log -Append -InputObject "Account Created for $name  | $Date "
		write "Account Created for $name"
		return "Created"
	}
	catch{
		$Status = "Failed to create account"
		$User.AD_Progress = "FAILED"
		write "Error creating account for $name, see log.txt for error report."
		SendEmail_Error "Unexpected error when creating account for $firstname $lastname"

		Out-File $PS_Log -Append -InputObject "Error: Creating account for $name  | $Date "
		Out-File $PS_Log -Append -InputObject $Error[0]
	}
}

function AddAddress([string]$Username){

	Set-ADUser `
	-Identity 			$Username `
	-StreetAddress 		"Street" `
	-City 				"City" `
	-Country 			"Country" `
	-PostalCode 		"0000" `
	-State 				"STATE"

}



function SetADGroups(){

		switch($JobTitle)
		{
			"Job1"{$ADArray 	= "Main_Group", "Group1"; break}
			"Job2"{$ADArray 	= "Main_Group", "Group2"; break}
			"Job3"{$ADArray 	= "Main_Group", "Group3"; break}
			
		}
		
	#Assign each group to the user
	for($i = 0; $i -lt $ADArray.Count; $i++){
		SetUserGroup $ADArray[$i]
	}

}


function SetUserGroup([string]$UserGroupName){

	try{
		Add-ADGroupMember `
		-Members $Username `
		-Identity $UserGroupName
		write "User group $UserGroupName set to $Username "
	}catch{
		Out-File $PS_Log -Append -InputObject "Error: Assigning group ($UserGroupName) to $name  |  $Date "
	}
}


function AssignPhoneNum(){
	#Load CSV to an arrayList
	$PhoneNumbers = Import-Csv $PhoneNumsFile
	[System.Collections.ArrayList]$NumArray = $PhoneNumbers

	#Assign the top number and remove it from the arrayList
	$UserPhoneNumber = $PhoneNumbers[0].Numbers
	$NumArray.RemoveAt(0)

	$NumArray | Export-Csv -path $PhoneNumsFile -NoTypeInformation 
	return "$UserPhoneNumber"
}


function SendEmail_Error([string]$BodyText){

	Try {
		Send-MailMessage `
		-From "testemail@domain.com" `
		-Subject "[Auto-Mail] Error Creating User In Active Directory" `
		-To $SenderEmail `
		-Body "$BodyText `n Please resolve this issue with an appropriate Username." 

	}
	Catch {
		write "Error: Sending Email!"
		Out-File $PS_Log -Append -InputObject "Error: Sending Email to $SenderEmail | $Date "
		Out-File $PS_Log -Append -InputObject $Error[0]
	}

}


function SendEmail_UserCreated(){

	Try {
		Send-MailMessage `
		-From "testemail@domain.com" `
		-Subject "[Auto-Mail] Active Directory User Created" `
		-To $ManagerEMail `
		-Body "An active directory user has been created for $Name `n
		E-Mail: $Email
		Temp Password: $TempPassword
		Phone Extension: $UserPhoneNumber"
	}
	Catch {
		write "Error: Sending Email!"
		Out-File $PS_Log -Append -InputObject "Error: Sending Email to $SenderEmail | $Date "
		Out-File $PS_Log -Append -InputObject $Error[0]
		
	}
}


function SetUpDB(){
	#Create a Backup of DB
	Import-Csv $UserDataBase | Export-Csv -path $User_DB_BackUp -NoTypeInformation -Force
	write "Database backed up."
}


function ExportToDB(){
	write "Exporting data to database..."
	#Export data to the .csv DB

	foreach ($User in $ADUsers){
		$FileExists = Test-Path $UserDataBase
		If ($FileExists -eq $true) {
			[System.Collections.ArrayList]$UserArrayList = $ADUsers
			$UserArrayList | Export-Csv -path $UserDataBase -NoTypeInformation -Force
			$ExportedUser = $User.username
			Out-File $PS_Log -Append -InputObject "Exported $ExportedUser to Database | $Date "
		}
	}
}


##StartHere##
startfunc
