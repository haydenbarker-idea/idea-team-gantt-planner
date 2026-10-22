param([datetime]$Since)
$ol = New-Object -ComObject Outlook.Application
$inbox = $ol.GetNamespace("MAPI").GetDefaultFolder(6).Items; $inbox.Sort("[ReceivedTime]", $true)
@($inbox | Select-Object -First 40 | Where-Object { $_.Subject -like "Snapshot ready*" -and $_.ReceivedTime -gt $Since }) | ForEach-Object { $_.ReceivedTime.ToString("HH:mm:ss") }
