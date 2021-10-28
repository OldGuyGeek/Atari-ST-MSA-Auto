Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Please Select File"
$OpenFileDialog.InitialDirectory = $initialDirectory
$OpenFileDialog.filter = "All files (*.*)| *.*"
# Out-Null supresses the "OK" after selecting the file.
$OpenFileDialog.ShowDialog() | Out-Null
$Global:SelectedFile = $OpenFileDialog.FileName
Write-Output $Global:SelectedFile
Set-Variable -Name "msafile" -Value $Global:SelectedFile
Get-Variable msafile -ValueOnly
Write-Output $msafile

$msainput = New-Object psobject -Property @{
Invoke1 = "start"
MsaPrg = "H:\Atari\Window Stuff\MSA Converter 2.1\msa.exe"
ConvertType = "diskimg_to_hdisk"
DestinationFolder = "H:\Atari\16bit\Emulators\D Drive\Extracts"
}

Echo "$($msainput.Invoke1) ($msainput.MsaPrg) ($msainput.ConvertType) ($msainput.DestinationFolder)"