#Author Wyatt Maddox
#Written for Powershell 5.1
#Requires Active Directory Module 


#Check for previous list to compare to
$run_check = Test-Path ($PSScriptRoot + "\old_list.csv")
if ($run_check -eq $false) {
    Get-ADComputer -Filter "enabled -eq 'true'" -SearchBase "OU=SUBEXAMPLE,OU=EXAMPLE,DC=EXAMPLEDOMAIN,DC=COM" -Properties name | select name | 
        Export-Csv -Path ($PSScriptRoot + "\old_list.csv") -NoTypeInformation}

#Make the new list to compare for new or removed deivces
Get-ADComputer -Filter "enabled -eq 'true'" -SearchBase "OU=SUBEXAMPLE,OU=EXAMPLE,DC=EXAMPLEDOMAIN,DC=COM" -Properties name | select name | 
    Export-Csv -Path ($PSScriptRoot + "\current_list.csv") -NoTypeInformation

#Load the CSVs into variables
$current_list = import-csv -Path ($PSScriptRoot + "\current_list.csv")
$old_list = import-csv -Path ($PSScriptRoot + "\old_list.csv")

#Compare the two lists
Compare-Object  $old_list $current_list -Property name | Export-Csv ($PSScriptRoot + "\results.csv") -NoTypeInformation

#Load Results of compare into CSV
$results_list = import-csv -Path ($PSScriptRoot + "\results.csv")

#Format the results for html email
if ($results_list -eq $null){
    $results_list = "No Changes found"
    }
    Else{
($results_list | 
    ForEach-Object {
        $_.SideIndicator = $_.SideIndicator -replace '=>','Added Device' -replace '<=','Removed Device'
        })
$results_list = ($results_list | Select-Object (@{expression={$_.name}; label= "Device" },@{expression={$_.SideIndicator}; label= "Status" }) | ConvertTo-Html -Fragment)
$results_list = ($results_list -replace "<table>", "<table style='border: 1px solid black'>")
$results_list = ($results_list -replace "<th>", "<th style='border: 1px solid black'>")
$results_list = ($results_list -replace "<td>", "<td style='border: 1px solid black'>")
$results_list = ($results_list -replace "<td style='border: 1px solid black'>Removed Device</td>", "<td bgcolor='#ff0000'>Removed Device</td>")
$results_list = ($results_list -replace "<td style='border: 1px solid black'>Added Device</td>", "<td bgcolor='#008000'>Added Device</td>")
}


#Email Vars
$smtp = "REPLACE WITH SMTP ADDRESS"  
$to = "First Last <user@example.com>" 
$from = "New Devices<newdevices@example.com>" 
$subject = "New Devices " + (Get-Date -Format "MM-dd-yy")
$date = (Get-Date -Format "MM-dd-yy") 
$domain_name = (Get-ADDomain | select NetBiosName)
$domain_name = $domain_name.NetBiosName
$body = @"
<body>
<h1>Below are the new added or removed devices:</h1>
<p>Date: $date</p>
<p>Domain: $domain_name</p>
$results_list
</body>
"@

#Send email of results to emails listed above
Send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body $body -BodyAsHtml #-Attachments ( ($PSScriptRoot + "\results.csv"))

#Delete old_list
Remove-Item -Path ($PSScriptRoot + "\old_list.csv")

#Rename current_list to old_list
Rename-Item -Path ($PSScriptRoot + "\current_list.csv") -NewName "old_list.csv"
