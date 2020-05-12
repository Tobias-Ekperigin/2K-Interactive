<#	
	.NOTES
	===========================================================================
	 Created on:   	12/05/2020 18:37
	 Created by:   	Tobias Ekperigin
	 Filename:     	Scripting Evaluation
     Organization:  2k Interactive
     Version:       1.0.0 - Test - 12/05/2020
	===========================================================================
	.DESCRIPTION
Simple Test to show interaction with a CSV. 

#>

#---------------------------------------------------------[Initialisations]---------------------------------------------------------------
$InputPath  = "C:\temp\Users.csv"
$ExportPath = "C:\temp\(ScriptingEvaluation)(Tobias.Ekperigin)$(get-date -format '(dd.MMM.yyyy_hh.mm)').csv"
#-----------------------------------------------------------[Execution]-------------------------------------------------------------------

if(($InputPath -ne $null -or " ") -and (Test-Path $InputPath)){

$Users = Import-Csv -Path $InputPath -Delimiter "," -Encoding UTF8; 

#2 How Many Users are there? 
Write-Host "Imported CSV Contains:" -f White -NoNewline; Write-Host " ($($Users.Count)) " -f Green -NoNewline; Write-Host "Users" -f White; 

#3 What is the total Size of all mailboxes? 
[decimal]$totalsize = 0; $Users | ForEach-Object {$totalsize += $_.MailboxsizeGB}; 
Write-Host "The total size of all Mailboxes = " -f White -NoNewline; Write-Host "$($totalsize) GB" -f Cyan;

#4 How many accounts exist with non-identical EmailAddress/UserPrincipalName?   
$Unique = $Users | Where-Object {$_.EmailAddress -cnotmatch $_.UserPrincipalName};
Write-Host "Total Number of non-identical Users: " -f White -NoNewline; Write-Host "$($Unique.count)" -f Green; Write-Output $Unique;


#5 What is the total Size of all mailboxes at the site NYC? 
[int]$totalsize = 0; $Users | Where-Object {$_.Site -eq "NYC"} | ForEach-Object {$totalsize += $_.MailboxsizeGB}; 
Write-Host "The total size of all Mailboxes in NYC = " -f White -NoNewline; Write-Host "$($totalsize) GB" -f Cyan;

#6 How many Employees (AccountType: Employee) have mailboxes larger than 10 GB?
$ELM = $Users | Where-Object {$_.AccountType -eq "Employee" -and [int]$_.MailboxSizeGB -ge '10'};
Write-Host "The Number of Employees with mailboxes larger than 10GB = " -f White -NoNewline; Write-Host "$($ELM.Count)" -f Green;

#7 Provide a list of the top 10 users with EmailAddress @domain2.com in Site: NYC by mailbox size, descending.
$Users | Where-Object {$_.Site -eq "NYC" -and ($_.UserPrincipalName -split "@" | Select-Object -Last 1) -eq 'domain2.com'} |`
Sort-Object -Property MailboxSizeGB -Descending | Select-Object -First '10' | Format-Table -AutoSize;

#7 a.) Extract Usernames from UPN and parse into single string.
$sorted = $Users | Where-Object {$_.Site -eq "NYC" -and ($_.UserPrincipalName -split "@" | Select-Object -Last 1) -eq 'domain2.com'} |`
Sort-Object -Property MailboxSizeGB -Descending | Select-Object -First '10'; 

[string]$string = $null; $sorted | ForEach-Object {$string += $(($_.UserPrincipalName -split "@" | Select-Object -First 1 ) + " ")}; Write-Output $string



#8 Create a new CSV file that summarizes Sites, using the following headers: Site, TotalUserCount, EmployeeCount, ContractorCount, TotalMailboxSizeGB, AverageMailboxSizeGB

$Report = New-Object System.Collections.ArrayList($null)

$sites = $Users.site | Select-Object -Unique;
$sites | ForEach-Object{

$TS = $_ 
$site = $Users | Where-Object {$_.Site -eq "$TS"};

[decimal]$TMS = 0; $site | ForEach-Object{[decimal]$TMS += $_.MailboxSizeGB};
[decimal]$AMS = $($site.MailboxSizeGB | Measure-Object -Average).Average;

        $tmpobj = [pscustomobject][ordered]@{
        Site                 = $($_);
        TotalUserCount       = $($site.count);
        EmployeeCount        = $(($site | Where-Object {$_.AccountType -eq "Employee"}) | Measure-Object).Count;
        ContractorCount      = $(($site | Where-Object {$_.AccountType -eq "Contractor"}) | Measure-Object).Count;
        TotalMailboxSizeGB   = $([math]::Round($TMS,1));
        AverageMailboxSizeGB = $([math]::Round($AMS,1));
        }

        $outputobj = $tmpobj | ConvertTo-Json
        [void]($Report.Add($outputobj)) 
}

$Report | ConvertFrom-Json | Export-Csv -Path $exportpath -Delimiter "," -Encoding UTF8 -NoTypeInformation;
Write-host "CSV Exported to: " -NoNewline -f white; Write-Host "$ExportPath" -f Magenta;  

}else{Write-Host "Input Path is not correct; Please review the input path variable under 'Initalisations'" -f Red}














