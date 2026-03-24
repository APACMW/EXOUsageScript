# Job file
param (
    [Parameter(Mandatory = $true)][string]$inputFile,
    [Parameter(Mandatory = $true)][string]$outputFile
)
$certthumbprint = '15958E05E3E4C2E563CE9BC346B25A2D70867048';
$appId = "24850000-9472-46ea-a062-402b17c79a9e";
$org = "vjqg8.onmicrosoft.com";
$cert = get-item "cert:\localmachine\my\$certthumbprint";
Connect-ExchangeOnline -AppId $appId -Organization $org -Certificate $cert;
$recipients = @(import-csv $inputFile);
$recipients | foreach-object { 
    $upn = $_.PrimarySmtpAddress.tostring();
    $mbx = get-mailbox $upn | Select-Object exchangeguid, ProhibitSendReceiveQuota, RecoverableItemsQuota, ArchiveStatus, AutoExpandingArchiveEnabled;
    $mbStat = Get-MailboxStatistics -Identity $mbx.exchangeguid;
    $outputObj = "" | Select-Object UserPrincipalName, ProhibitSendReceiveQuota, RecoverableItemsQuota, TotalItemSize, TotalDeletedItemSize, ArchiveStatus,AutoExpandingArchiveEnabled, UsageRatio,RecoverableItemUsageRatio;
    $outputObj.UserPrincipalName = $upn;
    $outputObj.ProhibitSendReceiveQuota = $mbx.ProhibitSendReceiveQuota;
    $outputObj.RecoverableItemsQuota = $mbx.RecoverableItemsQuota;
    $outputObj.TotalItemSize = $mbStat.TotalItemSize;
    $outputObj.TotalDeletedItemSize = $mbStat.TotalDeletedItemSize;
    $outputObj.ArchiveStatus = $mbx.ArchiveStatus;
    $outputObj.AutoExpandingArchiveEnabled = $mbx.AutoExpandingArchiveEnabled;
    $outputObj.UsageRatio = 0.0;
    $outputObj.RecoverableItemUsageRatio = 0.0;
    $outputObj | Export-Csv $outputFile -Append -NoTypeInformation;
}
Write-Host "Completed processing. Output saved to $outputFile";
Disconnect-ExchangeOnline -Confirm:$false;