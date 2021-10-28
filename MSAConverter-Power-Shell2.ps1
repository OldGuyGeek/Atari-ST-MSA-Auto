Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Please Select File"
$OpenFileDialog.InitialDirectory = $initialDirectory
$OpenFileDialog.filter = "ST Disk|*.ST|MSA Disk|*.MSA|All files| *.*"
# Out-Null supresses the "OK" after selecting the file.
$OpenFileDialog.ShowDialog() | Out-Null
$Global:SelectedFile = $OpenFileDialog.FileName
<#Write-Host $Global:SelectedFile#>
Set-Variable -Name "msafile" -Value $Global:SelectedFile
Get-Variable msafile -ValueOnly
<#Write-Host $msafile#>

$msainput = New-Object psobject -Property @{
NeededQuotes = """"
SingleQuotes ="'"
Invoke1 = "start"
MsaPrg = "H:\Atari\Window Stuff\MSA Converter 2.1\msa.exe"
ConvertType = "diskimg_to_hdisk"
DestinationFolder = "H:\Atari\16bit\Emulators\D Drive\Extracts"
}

$SubFolderName = [System.IO.Path]::GetFileNameWithoutExtension($Global:SelectedFile)
Write-Host $SubFolderName
Set-Location -Path $msainput.DestinationFolder
New-Item -Path $SubFolderName -ItemType Directory
$FullDestinationFolder = $msainput.DestinationFolder + "\" + $SubFolderName
write-host $FullDestinationFolder

<#Set-Variable -Name "NewDestinationFolder" -Value $FullDestinationFolder#>
Write-Host $NewDestinationFolder

Set-Location -Path "H:\Atari\Window Stuff\MSA Converter 2.1\"


Start-Process .\msa.exe "$($msainput.NeededQuotes)$($msainput.ConvertType)$($msainput.NeededQuotes) $($msainput.NeededQuotes)$msafile$($msainput.NeededQuotes) $($msainput.NeededQuotes)$($FullDestinationFolder)$($msainput.NeededQuotes)"
Start-Sleep -Seconds 10
Stop-Process -processname msa
<#Write-Host "$($msainput.MsaPrg) $($msainput.NeededQuotes)$($msainput.ConvertType)$($msainput.NeededQuotes) $($msainput.NeededQuotes)$msafile$($msainput.NeededQuotes) $($msainput.NeededQuotes)$($msainput.DestinationFolder)$($msainput.NeededQuotes)"
#>