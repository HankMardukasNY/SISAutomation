#File Paths
$csvpath="$PSScriptRoot\Student Information.csv"
$LogFile = "$PSScriptRoot\SISAutomation.log"
$NewStudents="$PSScriptRoot\NewStudents.csv"
$SISImport="$PSScriptRoot\SISImport.csv"

#OU Paths
$EastOUPath="OU=East_Students,OU=Students,OU=Contoso Accounts,DC=Contoso,DC=local"
$WestOUPath="OU=West_Students,OU=Students,OU=Contoso Accounts,DC=Contoso,DC=local"
$CentralOUPath="OU=Central_Students,OU=Students,OU=Contoso Accounts,DC=Contoso,DC=local"
$MemorialOUPath="OU=Memorial_Students,OU=Students,OU=Contoso Accounts,DC=Contoso,DC=local"
$StudentsOU = "OU=Students,OU=Contoso Accounts,DC=Contoso,DC=local"
$ArchiveOU = "OU=Graduates,OU=Archived,DC=Contoso,DC=local"
$HomeOUPath="OU=Home_Schooled,OU=Students,OU=Contoso Accounts,DC=Contoso,DC=local"

#School Groups
$EastGroups="DistStudents" , "O365StudentRestrict" , "365_License_StudentA5" , "PA-Filter-Students"
$WestGroups="DistStudents" , "O365StudentRestrict" , "365_License_StudentA5" , "PA-Filter-Students"
$CentralGroups="DistStudents" , "O365StudentRestrict" , "365_License_StudentA5" , "PA-Filter-Students"
$MemorialGroups="DistStudents" , "O365StudentRestrict" , "365_License_StudentA5" , "PA-Filter-Students"
$HomeGroups="DistStudents" , "O365StudentRestrict" , "PA-Filter-Students"
$ArchiveOU = "OU=Graduates,OU=Archived,DC=Contoso,DC=local"
$Date=(Get-Date -Format "MM/dd/yyyy")


#Delete last SIS import CSV

if (Test-Path $SISImport) {
  Remove-Item $SISImport
}
Add-Content -Path $SISImport -Value "ID Number,Import Type,School Name,Update Field,Update Value"

#Delete last SIS export CSV


if (Test-Path $csvpath) {
  Remove-Item $csvpath
}


#Update PNP Powershell
#Update-Module PnP.PowerShell

#Download latest SIS export from Contoso FTP
Start-Process -FilePath "C:\Program Files\CoreFTP\coreftp.exe" -ArgumentList '-s -O -site Contoso -d "/SISExport/Student Information.csv" -p "C:\ADScripts\"'


#Function to get current time
function Get-TimeStamp {
    
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date) }


#Connect to SharePoint Online
Connect-PnPOnline -Tenant Contoso.onmicrosoft.com -ClientId 3cb6adef-4639-4dec-bc2f-2353d6cbbe52 -Thumbprint 45245245245345345345345 -Url https://Contoso.sharepoint.com/sites/Q-Drive

#Download CSV files from SharePoint
Get-PnPFile -Url "Shared Documents/All Schools/New Students/NewStudents.csv" -Path $PSScriptRoot -FileName NewStudents.csv -AsFile -Force
Get-PnPFile -Url "Shared Documents/All Schools/New Students/SISAutomation.log" -Path $PSScriptRoot -FileName SISAutomation.log -AsFile -Force


function Upload-SharepointCSV {
#Upload the CSV files to SharePoint
$AddPNP=Add-PnPFile -Path $LogFile -Folder "Shared Documents/All Schools/New Students"
$AddPNP2=Add-PnPFile -Path $NewStudents -Folder "Shared Documents/All Schools/New Students"
}

function Upload-SIS {
#Upload new account emails to Contoso FTP
if (Test-Path $SISImport) {
Set-Content $SISImport ((Get-Content $SISImport) -replace '"')
Start-Process -FilePath "C:\Program Files\CoreFTP\coreftp.exe" -ArgumentList '-s -O -site Contoso -u "C:\ADScripts\SISImport.csv" -p "/SISExport/"'}
}

#Import Active Directory Module
Import-Module ActiveDirectory


#Import SIS List
Write-Output "$(Get-TimeStamp) Starting SIS Automation" | Out-file $LogFile -append
$SIS=Import-Csv -Path $csvpath

#SIS Student IDs
$SISStudentID=Import-Csv -Path $csvpath | select -ExpandProperty 'Student ID Number' 

#List all student AD accounts not in SIS
$Graduates = Get-AdUser -SearchBase $StudentsOU -Filter * -Properties * | Select SamAccountName, GivenName, Surname, Displayname, Description, UserPrincipalName, DistinguishedName, Company, Department, Office | Where { $SISStudentID -notcontains $_.Description }


#Function to remove students from AD that aren't in SIS Export
function Remove-Graduates {

#Disable accounts not in SIS
ForEach ($User in $Graduates){
$DisplayName = $User.DisplayName
Write-Warning "$DisplayName not in SIS, disabling"
Write-Output "$(Get-TimeStamp) $DisplayName not in SIS, disabling" | Out-file $LogFile -append
Set-ADUser -Identity $User.SamAccountName -Enabled $false
Move-ADObject -Identity $User.DistinguishedName -TargetPath $ArchiveOU }
}


#Adds new AD accounts. Moves AD accounts to correct school/group

function Add-Move-Students {

$SIS | ForEach-Object {

#CSV Variables
$FullName=$_."First Name"+" "+$_."Middle Name"+" "+$_."Last Name"
$DisplayName= [regex]::Replace($FullName, "\s+", " ")
$FirstName=$_."First Name"
$MiddleName=$_."Middle Name"
$LastName=$_."Last Name"
$StudentID=$_."Student ID Number"
$School=$_."Current School"
$Grade=$_.Grade


if($User=Get-ADUser -ldapfilter "(description=$StudentID)" -Property SamAccountName, GivenName, Surname, Displayname, Description, UserPrincipalName, DistinguishedName, Company, Department, Office, Enabled)
{
$SamAccountName=$User.SamAccountName
$Displayname=$User.DisplayName
$DistinguishedName=$User.DistinguishedName
$Company=$User.Company
$Department=$User.Department
if(!($Company)) {Set-ADUser -Identity $User -Company 'TempSchool'}
if(!($Department)) {Set-ADUser -Identity $SamAccountName -Department 'TempGrade'}


If ($User.enabled -eq $False){
Set-ADUser -Identity $SamAccountName -Enabled $True
Move-ADObject -Identity $DistinguishedName -TargetPath $StudentsOU
Write-Output "$(Get-TimeStamp) Re-Enabling $Displayname" | Out-file $LogFile -append }


if ($_.'Current School' -eq 'Contoso Central High School'){
            switch ($_.Grade) {
                "10" {
                    if(!($Department.Equals('CentralGrade10'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'CentralGrade10'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to CentralGrade10" | Out-file $LogFile -append }
                     }
                "11" {
                    if(!($Department.Equals('CentralGrade11'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'CentralGrade11'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to CentralGrade11" | Out-file $LogFile -append }
                     }
                "12" {
                    if(!($Department.Equals('CentralGrade12'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'CentralGrade12'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to CentralGrade12" | Out-file $LogFile -append }
                     }
               }
                if(!($Company.Equals("CentralStudents"))) {
                    Write-Output "$(Get-TimeStamp) Adding $Displayname to CentralStudents" | Out-file $LogFile -append
                    Set-ADUser -Identity $SamAccountName -Company 'CentralStudents'
                }
                if(!($DistinguishedName -like "*Central*")) {
                    Move-ADObject -Identity $DistinguishedName -TargetPath $CentralOUPath
                }
} elseif ($_.'Current School' -eq 'Contoso Memorial Junior High School'){
            switch ($_.Grade) {
                "6" {
                    if(!($Department.Equals('MemorialGrade6'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'MemorialGrade6'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to MemorialGrade6" | Out-file $LogFile -append }
                     }
                "7" {
                    if(!($Department.Equals('MemorialGrade7'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'MemorialGrade7'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to MemorialGrade7" | Out-file $LogFile -append }
                     }
                "8" {
                    if(!($Department.Equals('MemorialGrade8'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'MemorialGrade8'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to MemorialGrade8" | Out-file $LogFile -append }
                     }
                "9" {
                    if(!($Department.Equals('MemorialGrade9'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'MemorialGrade9'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to MemorialGrade9" | Out-file $LogFile -append }
                     }
               }
               if(!($Company.Equals("MemorialStudents"))) {
                    Write-Output "$(Get-TimeStamp) Adding $Displayname to MemorialStudents" | Out-file $LogFile -append
                    Set-ADUser -Identity $SamAccountName -Company 'MemorialStudents'
                }
                if(!($DistinguishedName -like "*Memorial*")) {
                    Move-ADObject -Identity $DistinguishedName -TargetPath $MemorialOUPath
                }
} elseif ($_.'Current School' -eq 'Contoso West High School'){
            switch ($_.Grade) {
                "6" {
                    if(!($Department.Equals('WestGrade6'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'WestGrade6'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to WestGrade6" | Out-file $LogFile -append }
                     }
                "7" {
                    if(!($Department.Equals('WestGrade7'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'WestGrade7'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to WestGrade7" | Out-file $LogFile -append }
                     }
                "8" {
                    if(!($Department.Equals('WestGrade8'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'WestGrade8'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to WestGrade8" | Out-file $LogFile -append }
                     }
                "9" {
                    if(!($Department.Equals('WestGrade9'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'WestGrade9'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to WestGrade9" | Out-file $LogFile -append }
                     }
                "10" {
                    if(!($Department.Equals('WestGrade10'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'WestGrade10'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to WestGrade10" | Out-file $LogFile -append }
                     }
                "11" {
                    if(!($Department.Equals('WestGrade11'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'WestGrade11'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to WestGrade11" | Out-file $LogFile -append }
                     }
                "12" {
                    if(!($Department.Equals('WestGrade12'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'WestGrade12'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to WestGrade12" | Out-file $LogFile -append }
                     }
               }
            if(!($Company.Equals("WestStudents"))) {
                Write-Output "$(Get-TimeStamp) Adding $Displayname to WestStudents" | Out-file $LogFile -append
                Set-ADUser -Identity $SamAccountName -Company 'WestStudents'
            }
                if(!($DistinguishedName -like "*West*")) {
                    Move-ADObject -Identity $DistinguishedName -TargetPath $WestOUPath
                }
} elseif ($_.'Current School' -eq 'Contoso East High School'){
            switch ($_.Grade) {
                "6" {
                    if(!($Department.Equals('EastGrade6'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'EastGrade6'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to EastGrade6" | Out-file $LogFile -append }
                     }
                "7" {
                    if(!($Department.Equals('EastGrade7'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'EastGrade7'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to EastGrade7" | Out-file $LogFile -append }
                     }
                "8" {
                    if(!($Department.Equals('EastGrade8'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'EastGrade8'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to EastGrade8" | Out-file $LogFile -append }
                     }
                "9" {
                    if(!($Department.Equals('EastGrade9'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'EastGrade9'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to EastGrade9" | Out-file $LogFile -append }
                     }
                "10" {
                    if(!($Department.Equals('EastGrade10'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'EastGrade10'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to EastGrade10" | Out-file $LogFile -append }
                     }
                "11" {
                    if(!($Department.Equals('EastGrade11'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'EastGrade11'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to EastGrade11" | Out-file $LogFile -append }
                     }
                "12" {
                    if(!($Department.Equals('EastGrade12'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'EastGrade12'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to EastGrade12" | Out-file $LogFile -append }
                     }
                "UGS" {
                    if(!($Department.Equals('UGS'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'UGS'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to UGS" | Out-file $LogFile -append }
                     }
               }
            if(!($Company.Equals("EastStudents"))) {
                Write-Output "$(Get-TimeStamp) Adding $Displayname to EastStudents" | Out-file $LogFile -append
                Set-ADUser -Identity $SamAccountName -Company 'EastStudents'
            }
                if(!($DistinguishedName -like "*East*")) {
                    Move-ADObject -Identity $DistinguishedName -TargetPath $EastOUPath
                }
} else{
            switch ($_.Grade) {
                "6" {
                    if(!($Department.Equals('HomeGrade7'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'HomeGrade6'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to HomeGrade6" | Out-file $LogFile -append }
                     }
                "7" {
                    if(!($Department.Equals('HomeGrade7'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'HomeGrade7'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to HomeGrade7" | Out-file $LogFile -append }
                     }
                "8" {
                    if(!($Department.Equals('HomeGrade8'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'HomeGrade8'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to HomeGrade8" | Out-file $LogFile -append }
                     }
                "9" {
                    if(!($Department.Equals('HomeGrade9'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'HomeGrade9'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to HomeGrade9" | Out-file $LogFile -append }
                     }
                "10" {
                    if(!($Department.Equals('HomeGrade10'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'HomeGrade10'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to HomeGrade10" | Out-file $LogFile -append }
                     }
                "11" {
                    if(!($Department.Equals('HomeGrade11'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'HomeGrade11'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to HomeGrade11" | Out-file $LogFile -append }
                     }
                "12" {
                    if(!($Department.Equals('HomeGrade12'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'HomeGrade12'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to HomeGrade12" | Out-file $LogFile -append }
                     }
                "UGS" {
                    if(!($Department.Equals('UGS'))) {
                        Set-ADUser -Identity $SamAccountName -Department 'UGS'
                        Write-Output "$(Get-TimeStamp) Adding $Displayname to UGS" | Out-file $LogFile -append }
                     }
               }
            if(!($Company.Equals("HomSISed"))) {
                Write-Output "$(Get-TimeStamp) Adding $Displayname to Home Schooled" | Out-file $LogFile -append
                Write-Host "$(Get-TimeStamp) Adding $Displayname to Home Schooled"         
                Set-ADUser -Identity $SamAccountName -Company 'HomSISed'
            }
                if(!($DistinguishedName -like "*Home*")) {
                    Move-ADObject -Identity $DistinguishedName -TargetPath $HomeOUPath
                }

}

} else {

Write-Warning "$DisplayName missing account, creating"
Write-Output "$(Get-TimeStamp) $DisplayName missing account, creating" | Out-file $LogFile -append

#Generates a random password from 100000 to 999999
$InputRange = 100000..999999
$Password = Get-Random -InputObject $InputRange

#Generates a username from first seven characters of last name and first inistal
$LastTrunkReplace=$LastName -replace "[:\-' .,/()]", ""
$LastTrunkLower=$LastTrunkReplace.ToLower() 
$LastTrunk=$LastTrunkLower.SubString(0, [System.Math]::Min(7, $LastTrunkLower.Length))
$FirstTrunkReplace=$FirstName -replace "[:\-' .,/()]", ""
$FirstTrunkLower=$FirstTrunkReplace.ToLower()
$FirstTrunk=$FirstTrunkLower.SubString(0, [System.Math]::Min(1, $FirstTrunkLower.Length))
$IDTrunkReplace=$StudentID -replace "[:\-' .,/()]", ""
$IDTrunk=$IDTrunkReplace.SubString($IDTrunkReplace.Length -4)
$SamAccountName=$LastTrunk + $FirstTrunk + $IDTrunk


# Find available username

if ($(Get-ADUser -Filter {SamAccountName -eq $SamAccountName})) {
   $i = 1
    do
    {   
        $suffix = "{0:d2}" -f $i
        $NewUsername = $SamAccountName + $suffix
        $i++ 
    } Until (!(Get-ADuser -Filter {SamAccountName -eq $NewUsername} -ErrorAction SilentlyContinue))
        $SamAccountName = $NewUsername
}


#Export result fields for Sharepoint
    $computerObject = [PSCustomObject]@{
        'First Name'          = $FirstName
        'Middle Name'         = $MiddleName
        'Last Name'           = $LastName
        'Student ID Number'   = $StudentID
        'Current School'      = $School
        'Grade'               = $Grade
        'Password'            = $Password
        'Email'               = ($SamAccountName  + "@Contoso.org")
        'Date Created'        = $Date
        }

#Export result fields for SIS
    $SISObject = [PSCustomObject]@{
        'ID Number'           = $StudentID
        'Import Type'         = 'Student'
        'School Name'         = $School
        'Update Field'        = '14'
        'Update Value'        = ($SamAccountName  + "@Contoso.org")
        }

#Create account based on school

if ($_.'Current School' -eq 'Contoso East High School'){
    New-ADUser -SamAccountName $SamAccountName -UserPrincipalName ($SamAccountName  + "@Contoso.org") -GivenName $_."First Name" -SurName $_."Last Name" –Description $_."Student ID Number" -Name $SamAccountName -DisplayName $DisplayName -Company 'EastStudents' -Path $EastOUPath -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) -Enabled $True -PasswordNeverExpires $True -CannotChangePassword $True -PassThru
    Add-ADPrincipalGroupMembership -Identity $SamAccountName -MemberOf $EastGroups
    Enable-ADAccount -Identity $SamAccountName
    $computerObject | Export-Csv -Path $NewStudents -NoTypeInformation -Append
    $SISObject | Export-Csv -Path $SISImport -NoTypeInformation -Append
    Write-Output "$(Get-TimeStamp) $SamAccountName has been created" | Out-file $LogFile -append
        switch ($_.Grade) {
            "6"  { Set-ADUser -Identity $samaccountname -Department 'EastGrade6' }
            "7"  { Set-ADUser -Identity $samaccountname -Department 'EastGrade7' }
            "8"  { Set-ADUser -Identity $samaccountname -Department 'EastGrade8' }
            "9"  { Set-ADUser -Identity $samaccountname -Department 'EastGrade9' }
            "10" { Set-ADUser -Identity $samaccountname -Department 'EastGrade10' }
            "11" { Set-ADUser -Identity $samaccountname -Department 'EastGrade11' }
            "12" { Set-ADUser -Identity $samaccountname -Department 'EastGrade12' }
            "UGS" { Set-ADUser -Identity $samaccountname -Department 'UGS' }
        }
}elseif ($_.'Current School' -eq 'Contoso West High School'){
    New-ADUser -SamAccountName $SamAccountName -UserPrincipalName ($SamAccountName  + "@Contoso.org") -GivenName $_."First Name" -SurName $_."Last Name" –Description $_."Student ID Number" -Name $SamAccountName -DisplayName $DisplayName -Company 'WestStudents' -Path $WestOUPath -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) -Enabled $True -PasswordNeverExpires $True -CannotChangePassword $True -PassThru
    Add-ADPrincipalGroupMembership -Identity $SamAccountName -MemberOf $WestGroups
    Enable-ADAccount -Identity $SamAccountName
    $computerObject | Export-Csv -Path $NewStudents -NoTypeInformation -Append
    $SISObject | Export-Csv -Path $SISImport -NoTypeInformation -Append
    Write-Output "$(Get-TimeStamp) $SamAccountName has been created" | Out-file $LogFile -append
        switch ($_.Grade) {
            "6"  { Set-ADUser -Identity $samaccountname -Department 'WestGrade6' }
            "7"  { Set-ADUser -Identity $samaccountname -Department 'WestGrade7' }
            "8"  { Set-ADUser -Identity $samaccountname -Department 'WestGrade8' }
            "9"  { Set-ADUser -Identity $samaccountname -Department 'WestGrade9' }
            "10" { Set-ADUser -Identity $samaccountname -Department 'WestGrade10' }
            "11" { Set-ADUser -Identity $samaccountname -Department 'WestGrade11' }
            "12" { Set-ADUser -Identity $samaccountname -Department 'WestGrade12' }
        }
}elseif ($_.'Current School' -eq 'Contoso Memorial Junior High School'){
    New-ADUser -SamAccountName $SamAccountName -UserPrincipalName ($SamAccountName  + "@Contoso.org") -GivenName $_."First Name" -SurName $_."Last Name" –Description $_."Student ID Number" -Name $SamAccountName -DisplayName $DisplayName -Company 'MemorialStudents' -Path $MemorialOUPath -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) -Enabled $True -PasswordNeverExpires $True -CannotChangePassword $True -PassThru
    Add-ADPrincipalGroupMembership -Identity $SamAccountName -MemberOf $MemorialGroups
    Enable-ADAccount -Identity $SamAccountName
    $computerObject | Export-Csv -Path $NewStudents -NoTypeInformation -Append
    $SISObject | Export-Csv -Path $SISImport -NoTypeInformation -Append
    Write-Output "$(Get-TimeStamp) $SamAccountName has been created" | Out-file $LogFile -append
        switch ($_.Grade) {
            "6"  { Set-ADUser -Identity $samaccountname -Department 'MemorialGrade6' }
            "7"  { Set-ADUser -Identity $samaccountname -Department 'MemorialGrade7' }
            "8"  { Set-ADUser -Identity $samaccountname -Department 'MemorialGrade8' }
            "9"  { Set-ADUser -Identity $samaccountname -Department 'MemorialGrade9' }
            "10" { Set-ADUser -Identity $samaccountname -Department 'MemorialGrade10' }
            "11" { Set-ADUser -Identity $samaccountname -Department 'MemorialGrade11' }
            "12" { Set-ADUser -Identity $samaccountname -Department 'MemorialGrade12' }
        }
}elseif ($_.'Current School' -eq 'Contoso Central High School'){
    New-ADUser -SamAccountName $SamAccountName -UserPrincipalName ($SamAccountName  + "@Contoso.org") -GivenName $_."First Name" -SurName $_."Last Name" –Description $_."Student ID Number" -Name $SamAccountName -DisplayName $DisplayName -Company 'CentralStudents' -Path $CentralOUPath -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) -Enabled $True -PasswordNeverExpires $True -CannotChangePassword $True -PassThru
    Add-ADPrincipalGroupMembership -Identity $SamAccountName -MemberOf $CentralGroups
    Enable-ADAccount -Identity $SamAccountName
    $computerObject | Export-Csv -Path $NewStudents -NoTypeInformation -Append
    $SISObject | Export-Csv -Path $SISImport -NoTypeInformation -Append
    Write-Output "$(Get-TimeStamp) $SamAccountName has been created" | Out-file $LogFile -append
        switch ($_.Grade) {
            "10" { Set-ADUser -Identity $samaccountname -Department 'CentralGrade10' }
            "11" { Set-ADUser -Identity $samaccountname -Department 'CentralGrade11' }
            "12" { Set-ADUser -Identity $samaccountname -Department 'CentralGrade12' }
        }
}else {
    New-ADUser -SamAccountName $SamAccountName -UserPrincipalName ($SamAccountName  + "@Contoso.org") -GivenName $_."First Name" -SurName $_."Last Name" –Description $_."Student ID Number" -Name $SamAccountName -DisplayName $DisplayName -Company 'HomSISed' -Path $HomeOUPath -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) -Enabled $True -PasswordNeverExpires $True -CannotChangePassword $True -PassThru 
    Add-ADPrincipalGroupMembership -Identity $SamAccountName -MemberOf $HomeGroups 
    Enable-ADAccount -Identity $SamAccountName 
    #$computerObject | Export-Csv -Path $NewStudents -NoTypeInformation -Append
    $SISObject | Export-Csv -Path $SISImport -NoTypeInformation -Append
    Write-Output "$(Get-TimeStamp) $SamAccountName has been created" | Out-file $LogFile -append
        switch ($_.Grade) {
            "6"  { Set-ADUser -Identity $samaccountname -Department 'HomeGrade6' }
            "7"  { Set-ADUser -Identity $samaccountname -Department 'HomeGrade7' }
            "8"  { Set-ADUser -Identity $samaccountname -Department 'HomeGrade8' }
            "9"  { Set-ADUser -Identity $samaccountname -Department 'HomeGrade9' }
            "10" { Set-ADUser -Identity $samaccountname -Department 'HomeGrade10' }
            "11" { Set-ADUser -Identity $samaccountname -Department 'HomeGrade11' }
            "12" { Set-ADUser -Identity $samaccountname -Department 'HomeGrade12' }
            "UGS" { Set-ADUser -Identity $samaccountname -Department 'UGS' }
        }
      }
   }
}
Write-Output "$(Get-TimeStamp) Ending SIS Automation" | Out-file $LogFile -append
}


Remove-Graduates
Add-Move-Students
Upload-SharepointCSV
Upload-SIS