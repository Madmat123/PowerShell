$ExportFolder = "C:\Users\mmied\OneDrive\Pulpit\Excel"
//Get-Process | Export-Excel -Path "$ExportFolder\Process.xlsx" -FreezeTopRow -AutoSize
$Users = Import-Excel -Path "$ExportFolder\Import.xlsx" -WorksheetName "Users" | Select-Object Name,Age,Country
$Countries = Import-Excel -Path "$ExportFolder\Import.xlsx" -WorksheetName "Countries" | Select-Object CountryCode,Country
$DetailedUsers = [System.Collections.ArrayList]@()
foreach($User in $Users){
$CountryName = ($Countries | Where-Object CountryCode -eq $User.Country).Country
$NewUser = New-Object -TypeName psobject
Add-Member -InputObject $NewUser -MemberType NoteProperty -Name "Name" -Value $User.Name
Add-Member -InputObject $NewUser -MemberType NoteProperty -Name "Age" -Value $User.Age
Add-Member -InputObject $NewUser -MemberType NoteProperty -Name "CountryCode" -Value $User.Country
Add-Member -InputObject $NewUser -MemberType NoteProperty -Name "Country" -Value $CountryName
[void]$DetailedUsers.Add($NewUser)
}

$DetailedUsers | Export-Excel -Path "$ExportFolder\Import.xlsx" -WorksheetName "Users w Country Detailed"