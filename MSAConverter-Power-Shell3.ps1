<# NOTES
1. Does not gracefully exit if file selection is cancelled
2. Does not check validity of .zip files
3. Does not restrict output directory name to Atari 8 Character Limit
4. Does not extract .stx (MSA limitation)
#>


# Create System File Open Diaglogue Box
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Please Select File"
$OpenFileDialog.InitialDirectory = $initialDirectory

# Create filter for allowable file types
$OpenFileDialog.filter = "ST Disk|*.st|MSA Disk|*.MSA|Zips|*.zip"

# Out-Null supresses the "OK" after selecting the file.
$OpenFileDialog.ShowDialog() | Out-Null
$Global:SelectedFile = $OpenFileDialog.FileName

# Put the name of the selected file in a variable
Set-Variable -Name "msafile" -Value $Global:SelectedFile
Get-Variable msafile -ValueOnly

#Create object to hold all the needed variables Easier for later use
$msainput = New-Object psobject -Property @{
NeededQuotes = """"
SingleQuotes ="'"
# MsaPrgLoc = "H:\Atari\Window Stuff\MSA Converter 2.1\msa.exe"
ConvertType = "diskimg_to_hdisk"
DestinationFolder = "H:\Atari\16bit\Emulators\D Drive\Extracts"
}

# To get destination subfolder name based on the name of the file selected
$SubFolderName = [System.IO.Path]::GetFileNameWithoutExtension($Global:SelectedFile)

# Going to the Destination folder in order to create new directory
Set-Location -Path $msainput.DestinationFolder


# Create new directory based on selected file
New-Item -Path $SubFolderName -ItemType Directory

# Contatating names into final destination 
$FullDestinationFolder = $msainput.DestinationFolder + "\" + $SubFolderName

# Going to MSA location to invoke it
Set-Location -Path "H:\Atari\Window Stuff\MSA Converter 2.1\"

<# NOTE! NO CHECKING IF DIRECTORY ALREADY EXISTS #>

# Starting MSA with command line parameters
Start-Process .\msa.exe "$($msainput.NeededQuotes)$($msainput.ConvertType)$($msainput.NeededQuotes) $($msainput.NeededQuotes)$msafile$($msainput.NeededQuotes) $($msainput.NeededQuotes)$($FullDestinationFolder)$($msainput.NeededQuotes)"

# Lazy way to make sure files are finished writing before closing MSA
Start-Sleep -Seconds 10

# Closing MSA
Stop-Process -processname msa
