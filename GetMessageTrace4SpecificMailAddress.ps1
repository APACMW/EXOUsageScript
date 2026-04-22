<#
.SYNOPSIS
    This script retrieves message trace details for a specified email address within a given date range.
.PARAMETER queryType
    The type of query to perform. Valid values are 'sender' or 'recipient'. 
.PARAMETER emailAddress
    The email address to query for message traces.
.PARAMETER startDate
    The start date for the message trace query in 'yyyy-MM-dd' format.
.PARAMETER endDate
    The end date for the message trace query in 'yyyy-MM-dd' format.
.EXAMPLE
    .\GetMessageTrace4SpecificMailAddress.ps1 -queryType sender -emailAddress freeman@vjqg8.onmicrosoft.com -startDate '2026-04-14' -endDate '2026-04-16';
#>
param(
  [Parameter(Mandatory = $true)]
  [ValidateSet('sender', 'recipient')]
  [string]$queryType,

  [Parameter(Mandatory = $true)]
  [string]$emailAddress,

  [Parameter(Mandatory = $true)]
  [ValidatePattern('^\d{4}-\d{2}-\d{2}$')]
  [string]$startDate,

  [Parameter(Mandatory = $true)]
  [ValidatePattern('^\d{4}-\d{2}-\d{2}$')]
  [string]$endDate
)

class MessageTrace {  
  [string]$messageId;
  [string]$subject;
  [string]$senderAddress;
  [string]$recipientAddress;
  [datetime]$receivedDateTime;
  [string]$status;
  [string]$messageTraceID;

  [int] GetHashCode() {
    Write-Host "GetHashCode method";
    $combinedString = "$($this.messageId)|$($this.subject)|$($this.senderAddress)|$($this.recipientAddress)|$($this.receivedDateTime)|$($this.status)|$($this.messageTraceID)"
    return $combinedString.GetHashCode()
  }

  [bool] Equals([object]$other) {
    Write-Host "Equals method";
    if ($other -isnot [MessageTrace]) { 
      return $false 
    }
    return ($this.messageId -eq $other.messageId) -and
    ($this.subject -eq $other.subject) -and
    ($this.senderAddress -eq $other.senderAddress) -and
    ($this.recipientAddress -eq $other.recipientAddress) -and
    ($this.receivedDateTime -eq $other.receivedDateTime) -and
    ($this.status -eq $other.status) -and
    ($this.messageTraceID -eq $other.messageTraceID)
  }
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path;
Set-Location -Path $scriptDir;
Start-Transcript -Path ".\Transcript-$(Get-Date -Format 'yyyyMMddHHmmss').txt" -Append;
$outputFile = ".\MessageTraces-$(Get-Date -Format 'yyyyMMddHHmmss').csv";

$startDateTime = [datetime]"$($startDate)T00:00:00Z";
$endDateTime = [datetime]"$($endDate)T00:00:00Z";
Connect-ExchangeOnline;
$resultSize = 5000;
$sleepSeconds = 1;
$i = 1;
$msgTraces = New-Object 'System.Collections.Generic.List[MessageTrace]';

while ($startDateTime -lt $endDateTime) {
  Write-Host "Step $i for $emailAddress" -BackgroundColor Green -ForegroundColor White;
  $startDateTimeString = $startDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ");
  $pendDatetime = if ($startDateTime.AddHours(1) -gt $endDateTime) { $endDateTime } else { $startDateTime.AddHours(1) }
  $endDateTimeString = $pendDatetime.ToString("yyyy-MM-ddTHH:mm:ssZ");
  $traceParams = @{
    StartDate  = $startDateTimeString
    EndDate    = $endDateTimeString
    ResultSize = $resultSize
  };
  if ($queryType -eq 'sender') {
    $traceParams['SenderAddress'] = $emailAddress
  }
  else {
    $traceParams['RecipientAddress'] = $emailAddress
  }
  $records = @(Get-MessageTraceV2 @traceParams | Select-Object -Unique Received, SenderAddress, RecipientAddress, messageID, messageTraceID, Status, Subject); 
  $records | ForEach-Object {
    $msgTrace = [MessageTrace]::new();
    $msgTrace.messageId = $_.MessageID;
    $msgTrace.subject = $_.Subject;
    $msgTrace.senderAddress = $_.SenderAddress;
    $msgTrace.recipientAddress = $_.RecipientAddress;
    $msgTrace.receivedDateTime = $_.Received;
    $msgTrace.status = $_.Status;
    $msgTrace.messageTraceID = $_.MessageTraceID;
    if (!$msgTraces.Contains($msgTrace)) {
      $msgTraces.Add($msgTrace);
    }
  };     
  $i++;
  Start-Sleep -Seconds $sleepSeconds;
  $startDateTime = $startDateTime.AddHours(1);
}
$msgTraces | Export-Csv -Path $outputFile -NoTypeInformation -Append;
Disconnect-ExchangeOnline -Confirm:$false;
Stop-Transcript;