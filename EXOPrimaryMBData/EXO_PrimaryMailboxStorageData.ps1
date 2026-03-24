
<#
.SYNOPSIS
    This script retrieves primary mailbox storage data for all user mailboxes in an Exchange Online organization.
    Please configure the certificate-based authentication for EXO connection. Install the certificate under "cert:\localmachine\my\(https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps)
.DESCRIPTION
    This script retrieves primary mailbox storage data for all user mailboxes in an Exchange Online organization. 
    It connects to Exchange Online using certificate-based authentication, retrieves the list of user mailboxes, and then processes them in batches using background jobs. 
    The script collects storage data such as TotalItemSize, TotalDeletedItemSize, ProhibitSendReceiveQuota, and RecoverableItemsQuota for each mailbox. Finally, it calculates the usage ratios and outputs the results to a CSV file.
.PARAMETER thresholdRatio 
    The threshold ratio for mailbox storage usage. Mailboxes exceeding this ratio will be flagged.
.PARAMETER batchSize
    The number of mailboxes to process in each batch. Default is 50.
.EXAMPLE
    .\EXO_PrimaryMailboxStorageData.ps1
.NOTES
    author:Qi Dong (doqi@microsoft.com)
#>
[CmdletBinding()]
Param (
    [Parameter(Position = 0, Mandatory = $false)]    
    [double]$thresholdRatio = 0.8,
    [Parameter(Position = 1, Mandatory = $false)]
    [int] $batchSize = 50
)
function Get-Bytes {
    param (
        [string]$size
    )
    if ($size.Contains('(') -and $size.Contains(')')) {
        $str = $size.Substring($size.IndexOf('(') + 1, $size.IndexOf(')') - $size.IndexOf('(') - 1);
        $str = $str.Replace('bytes', '')
        $str = $str.Replace(',', '')
        $str = $str.Replace(' ', '')
    }
    $str;
}

# set the current location to the script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path;
Set-Location -Path $scriptDir;

# prepare recipient list
$RecipientFilePath = [System.IO.Path]::Combine($scriptDir, 'recipients.csv');
if (Test-Path -path $RecipientFilePath) {
    Remove-Item -Path $RecipientFilePath;
}
$certthumbprint = '15958E05E3E4C2E563CE9BC346B25A2D70867048';
$appId = "24850000-9472-46ea-a062-402b17c79a9e";
$org = "vjqg8.onmicrosoft.com";
$cert = get-item "cert:\localmachine\my\$certthumbprint";
Connect-ExchangeOnline -AppId $appId -Organization $org -Certificate $cert;
$data = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize unlimited;
$data | Select-Object -Property "DisplayName", "PrimarySmtpAddress", "Identity" | Export-Csv -Path $RecipientFilePath -NoTypeInformation;
Disconnect-ExchangeOnline -Confirm:$false;

# process recipients in batches using jobs
$recipients = @(import-csv $RecipientFilePath);
$outputFolder = $scriptDir;
$RecipientDataPath = [System.IO.Path]::Combine($outputFolder, 'RecipientData');
if (-not (Test-Path -Path $RecipientDataPath)) {
    New-Item -ItemType Directory -Path $RecipientDataPath | Out-Null;
}
$recipientFiles = Get-ChildItem -Path $RecipientDataPath -File -Filter '*.csv';
$recipientFiles | ForEach-Object { Remove-Item -Path $PSItem.FullName };
$outputFile = [System.IO.Path]::Combine($scriptDir, 'stats.csv');
$finalOutputFile = [System.IO.Path]::Combine($scriptDir, 'finalStats.csv');
if ([System.IO.File]::Exists($outputFile)) {
    Remove-Item -Path $outputFile;    
}
if ([System.IO.File]::Exists($finalOutputFile)) {
    Remove-Item -Path $finalOutputFile;    
}

$total = 0;
$pageIndex = 0; 
while ($total -lt $recipients.Count) {
    $pageIndex = $pageIndex + 1;
    $batch = $recipients | Select-Object -Skip $total -First $batchSize;
    $recoveryKeyFile = [System.IO.Path]::Combine($RecipientDataPath, ($pageIndex.ToString() + '.csv'));
    $batch | Export-Csv -Path $recoveryKeyFile -NoTypeInformation;
    $total += $batchSize;
}

$recipientFiles = Get-ChildItem -Path $RecipientDataPath -File -Filter '*.csv';
$jobIndex = 0;
$recipientFiles | ForEach-Object { 
    $inputFile = $PSItem.FullName;
    $scriptFile = [System.IO.Path]::Combine($scriptDir, 'job.ps1');
    $job = Start-Job -Name "Job_$jobIndex" -filepath "$scriptFile" -argumentlist $inputFile, $outputFile;
    $job | Wait-Job
    $job | Remove-Job;
    $jobIndex++;
    Remove-Item -Path $inputFile;
}

$records = Import-Csv -Path $outputFile;
$records | ForEach-Object {
    $r = $psitem;
    $r.ProhibitSendReceiveQuota = Get-Bytes $r.ProhibitSendReceiveQuota;
    $r.RecoverableItemsQuota = Get-Bytes $r.RecoverableItemsQuota;
    $r.TotalItemSize = Get-Bytes $r.TotalItemSize;
    $r.TotalDeletedItemSize = Get-Bytes $r.TotalDeletedItemSize;
    if ($r.ProhibitSendReceiveQuota -gt 0) {
        $r.UsageRatio = [System.Math]::Min(1.0, $r.TotalItemSize * 1.0 / $r.ProhibitSendReceiveQuota);
    }
    else {
        $r.UsageRatio = 0.0;
    }
    if ($r.RecoverableItemsQuota -gt 0) {
        $r.RecoverableItemUsageRatio = [System.Math]::Min(1.0, $r.TotalDeletedItemSize * 1.0 / $r.RecoverableItemsQuota);
    }
    else {
        $r.RecoverableItemUsageRatio = 0.0;
    }
}
Remove-Item -Path $outputFile;
$records | Export-Csv $outputFile -Append -NoTypeInformation;
$records | Where-Object { $_.UsageRatio -ge $thresholdRatio -or $_.RecoverableItemUsageRatio -ge $thresholdRatio } | Export-Csv $finalOutputFile -Append -NoTypeInformation;
Write-Host "All done. Final output saved to $finalOutputFile";
