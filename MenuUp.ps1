
### MenuUploader v1.0.0
### This program allows end-users to easily upload the Cafe Menu to Sharepoint
### Jarred Reid 2019

### Changelog:
### Version 1.0: Initial release

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '216,192'
$Form.text                       = "MenuUp"
$Form.TopMost                    = $false

#This base64 string holds the bytes that make up the company icon
$iconBase64      = '
$stream          = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$iconImage       = [System.Drawing.Image]::FromStream($stream, $true)
$Form.Icon       = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())

$uploadButton                    = New-Object system.Windows.Forms.Button
$uploadButton.text               = "Upload"
$uploadButton.width              = 140
$uploadButton.height             = 30
$uploadButton.location           = New-Object System.Drawing.Point(35,132)
$uploadButton.Font               = 'Microsoft Sans Serif,10,style=Bold'

$browseButton                    = New-Object system.Windows.Forms.Button
$browseButton.text               = "Browse"
$browseButton.width              = 140
$browseButton.height             = 30
$browseButton.location           = New-Object System.Drawing.Point(34,44)
$browseButton.Font               = 'Microsoft Sans Serif,10,style=Bold'

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Select the new menu"
$Label1.AutoSize                 = $true
$Label1.width                    = 140
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(43,23)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Upload the new menu"
$Label2.AutoSize                 = $true
$Label2.width                    = 140
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(37,110)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$Form.controls.AddRange(@($uploadButton,$browseButton,$Label1,$Label2))

 ### START BROWSE BUTTON
 $browseButton.Add_MouseClick({  
     
if($browseButton.Add_Click){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Company\_Workstation"
    $OpenFileDialog.filter = "PDF files (.pdf)| *.pdf*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
    $Global:SelectedFile = $OpenFileDialog.filename
    }
})
    ### END BROWSE BUTTON

    $uploadButton.Add_MouseClick({
        if($uploadButton.Add_Click){Process-File}
        })
    
Function Process-File{

    try{
        Copy-Item -Path $Global:SelectedFile -Destination "C:\Users\$env:USERNAME\Company Sharepoint Library\Cafe Menu.pdf" -Force

        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $MessageBody = "Menu uploaded successfully."
        $MessageTitle = "Upload Status"
         
        $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle)
    }

    catch{

        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $MessageBody = "Upload failed. Please try again."
        $MessageTitle = "Upload Status"
         
        $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle)

        $ErrorMessage= $._Exception.MessageBox
        Write-Host $ErrorMessage
    }
}
[void]$Form.ShowDialog()
