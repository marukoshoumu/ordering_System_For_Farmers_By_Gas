param(
    [Parameter(Mandatory=$true)][string]$FilePath,
    [string]$PrinterName
)

if ($PrinterName) {
    Start-Process -FilePath $FilePath -Verb PrintTo -ArgumentList $PrinterName -WindowStyle Hidden -Wait
} else {
    Start-Process -FilePath $FilePath -Verb Print -WindowStyle Hidden -Wait
}
