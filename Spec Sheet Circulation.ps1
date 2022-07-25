### Spec Sheet Circulation - Revised 7.21.2022

#region-----Create Environment

#region      Create Error Pop-Ups Function
Function New-Popup {

<#
.Synopsis
Display a Popup Message
.Description
This command uses the Wscript.Shell PopUp method to display a graphical message
box. You can customize its appearance of icons and buttons. By default the user
must click a button to dismiss but you can set a timeout value in seconds to 
automatically dismiss the popup. 

The command will write the return value of the clicked button to the pipeline:
  OK     = 1
  Cancel = 2
  Abort  = 3
  Retry  = 4
  Ignore = 5
  Yes    = 6
  No     = 7

If no button is clicked, the return value is -1.
.Example
PS C:\> new-popup -message "The update script has completed" -title "Finished" -time 5

This will display a popup message using the default OK button and default 
Information icon. The popup will automatically dismiss after 5 seconds.
.Notes
Last Updated: April 8, 2013
Version     : 1.0

.Inputs
None
.Outputs
integer

Null   = -1
OK     = 1
Cancel = 2
Abort  = 3
Retry  = 4
Ignore = 5
Yes    = 6
No     = 7
#>

Param (
[Parameter(Position=0,Mandatory=$True,HelpMessage="Enter a message for the popup")]
[ValidateNotNullorEmpty()]
[string]$Message,
[Parameter(Position=1,Mandatory=$True,HelpMessage="Enter a title for the popup")]
[ValidateNotNullorEmpty()]
[string]$Title,
[Parameter(Position=2,HelpMessage="How many seconds to display? Use 0 require a button click.")]
[ValidateScript({$_ -ge 0})]
[int]$Time=0,
[Parameter(Position=3,HelpMessage="Enter a button group")]
[ValidateNotNullorEmpty()]
[ValidateSet("OK","OKCancel","AbortRetryIgnore","YesNo","YesNoCancel","RetryCancel")]
[string]$Buttons="OK",
[Parameter(Position=4,HelpMessage="Enter an icon set")]
[ValidateNotNullorEmpty()]
[ValidateSet("Stop","Question","Exclamation","Information" )]
[string]$Icon="Information"
)

#convert buttons to their integer equivalents
Switch ($Buttons) {
    "OK"               {$ButtonValue = 0}
    "OKCancel"         {$ButtonValue = 1}
    "AbortRetryIgnore" {$ButtonValue = 2}
    "YesNo"            {$ButtonValue = 4}
    "YesNoCancel"      {$ButtonValue = 3}
    "RetryCancel"      {$ButtonValue = 5}
}

#set an integer value for Icon type
Switch ($Icon) {
    "Stop"        {$iconValue = 16}
    "Question"    {$iconValue = 32}
    "Exclamation" {$iconValue = 48}
    "Information" {$iconValue = 64}
}

#create the COM Object
Try {
    $wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
    #Button and icon type values are added together to create an integer value
    $wshell.Popup($Message,$Time,$Title,$ButtonValue+$iconValue)
}
Catch {
    #You should never really run into an exception in normal usage
    Write-Warning "Failed to create Wscript.Shell COM object"
    Write-Warning $_.exception.message
}

} #end function
#endregion

#region      Create Notice before executing

New-Popup -message "Please ensure the pending revision is located in the LITERATURE PENDING folder." -title "Notice" -time 0 -Icon Information

#endregion

#region      Create Product Line Input Box
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Spec Sheet Circulation'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Please enter the Product Line:'
$form.Controls.Add($label)

$labelPLex = New-Object System.Windows.Forms.Label
$labelPLex.Location = New-Object System.Drawing.Point(10,85)
$labelPLex.Size = New-Object System.Drawing.Size(280,20)
$labelPLex.Text = '(e.g. QX1, QX2, QF1, QF2, NEXUS, RP1)'
$form.Controls.Add($labelPLex)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)

$form.Topmost = $true

$form.Add_Shown({$textBox.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $pline = $textBox.Text
    $pline
}

if ($pline.Length -lt 2 -or $pline.Length -gt 5 -or $pline -notmatch "[A-Z][0-9]")
{
    New-Popup "Entered product line name is invalid. The program will now terminate. Make sure you are using the appropriate product line format. (Example: QX2, QF1, QXM, RP1)" -Title "Error!" -Buttons OK -Icon Exclamation -Time 20
    throw "Product Line invalid"
}

#endregion

#region      Create Part Number Input Box
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Spec Sheet Circulation'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(10,20)
$label1.Size = New-Object System.Drawing.Size(280,20)
$label1.Text = 'Please enter the Part Number:'
$form.Controls.Add($label1)

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(10,40)
$textBox1.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox1)

$form.Topmost = $true

$form.Add_Shown({$textBox1.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $pnumber = $textBox1.Text
    $pnumber
}

if ($pnumber.Length -lt 6 -or $pnumber.Length -gt 35)
{
    new-popup "Entered part number is invalid. The program will now terminate." -Title "Error!" -Buttons OK -Icon Exclamation -Time 10
    throw "Part Number invalid"
}
#endregion

#region      Create Old Rev Level Input Box
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Spec Sheet Circulation'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(10,20)
$label3.Size = New-Object System.Drawing.Size(280,20)
$label3.Text = 'Please enter the old Rev level:'
$form.Controls.Add($label3)

$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(10,40)
$textBox3.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox3)

$form.Topmost = $true

$form.Add_Shown({$textBox3.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $revOld = $textBox3.Text
    $revOld
}

if ($revOld.Length -lt 1 -or $revOld.Length -gt 1 -or $revOld -cnotmatch "[A-Z]")
{
    new-popup "Entered revision level is invalid. The program will now terminate." -Title "Error!" -Buttons OK -Icon Exclamation
    throw "Old Revision Level invalid"
}
#endregion

#region      Create New Rev Level Input Box
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Spec Sheet Circulation'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,20)
$label2.Size = New-Object System.Drawing.Size(280,20)
$label2.Text = 'Please enter the new Rev level:'
$form.Controls.Add($label2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10,40)
$textBox2.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox2)

$form.Topmost = $true

$form.Add_Shown({$textBox2.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $revNew = $textBox2.Text
    $revNew
}

if ($revNew.Length -lt 1 -or $revNew.Length -gt 1 -or $revNew -cnotmatch "[A-Z]")
{
new-popup "Entered revision level is invalid. The program will now terminate." -Title "Error!" -Buttons OK -Icon Exclamation
throw "New Revision Level invalid"
}

#endregion

#region      Create ConvertWordTo-PDF Function
function ConvertWordTo-PDF {
 
<# 
  
.SYNOPSIS 
 
ConvertTo-PDF converts Microsoft Word documents to PDF files. 
  
.DESCRIPTION 
 
The cmdlet queries the given source folder including sub-folders to find *.docx and *.doc files, 
converts all found files and saves them as pdf in the Destination folder. After completition, the Destination
folder with the newly created PDF files will be opened with Windows Explorer.
  
.PARAMETER SourceFolder
  
Mandatory. Enter the source folder of your Microsoft Word documents.
  
.PARAMETER DestinationFolder
 
Optional. Enter the Destination folder to save the created PDF documents. If you omit this parameter, pdf files will
be saved in the Source Folder.
 
.EXAMPLE 
 
ConvertWordTo-PDF -SourceFolder C:\Temp -DestinationFolder C:\Temp1
ConvertWordTo-PDF -SourceFolder C:\temp
  
.NOTES 
Author: Patrick Gruenauer | Microsoft PowerShell MVP [2018-2021] 
Web: https://sid-500.com 
  
#>
 
[CmdletBinding()]
 
param
(
  
[Parameter (Mandatory=$true,Position=0)]
[String]
$SourceFolder,
  
[Parameter (Position=1)]
[String]
$DestinationFolder = $SourceFolder
 
)
 
    $i = 0
 
    $word = New-Object -ComObject word.application 
    $FormatPDF = 17
    $word.visible = $false
    $types = '*.docx','*.doc'
 
    If ((Test-Path $SourceFolder) -eq $false) {
     
    throw "Error. Source Folder $SourceFolder not found." } 
 
    If ((Test-Path $DestinationFolder) -eq $false) {
     
    throw "Error. Destination Folder $DestinationFolder not found." } 
     
    $files = Get-ChildItem -Path $SourceFolder -Include $Types -Recurse -ErrorAction Stop
    ''
    Write-Warning "Converting Files to PDF ..."
    ''
     
    foreach ($f in $files) {
 
        $path = $DestinationFolder + '\' + $f.Name.Substring(0,($f.Name.LastIndexOf('.')))
        $doc = $word.documents.open($f.FullName) 
        $doc.saveas($path,$FormatPDF) 
        $doc.close()
        Write-Output "$($f.Name)"
        $i++
 
    }
    ''
    Write-Output "$i file(s) converted."
    Start-Sleep -Seconds 2 
    Invoke-Item $DestinationFolder
    $word.Quit()
     
     
}
#endregion

#endregion

#region-----Perform Tasks

#Create New Archive folder in case of Rev A archival
if ($revOld -match "\bA\b") 
    {
    New-Item -Path "N:\Deep_Archive_Documentation\DOCS\Literature\Product Spec Sheets\$pline\$pnumber" -ItemType Directory
    }
#" '\bA\b' requires the string to be an exact match, meaning it cannot be "AA" or "AB" or "AC", etc.

#region #Check for Files before Executing remainder of script - If no file in original 3 folders, terminate program.

$cusLitFileTest = Test-Path -Path "F:\Documentation\DOCS\Customer Literature\Spec Sheets\$pline\$pnumber Rev $revOld.pdf" -PathType Leaf
$litPenFileTest = Test-Path -Path "F:\Documentation Pending\Literature PENDING\$pnumber Rev $revNew.docx" -PathType Leaf
$litSpecSheetFileTest = Test-Path -Path "F:\Documentation\DOCS\Literature\Product Spec Sheets\$pline\$pnumber.docx" -PathType Leaf

$terminateNoFileFound = if (($litPenFileTest -or $cusLitFileTest -or $litSpecSheetFileTest) -eq $false)
    {
    New-Popup "One or more files not found" -Title "Error" -Buttons OK -Icon Exclamation -Time 20
    throw "One or more files not found."
    }

#endregion

#Move Old Revision .docx to Archives
Move-Item -Path "F:\Documentation\DOCS\Literature\Product Spec Sheets\$pline\$pnumber.docx" -Destination "N:\Deep_Archive_Documentation\DOCS\Literature\Product Spec Sheets\$pline\$pnumber\$pnumber Rev $revOld.docx" -Force
#

#Delete Old Revision .pdf file from circulation
Remove-Item -Path "F:\Documentation\DOCS\Customer Literature\Spec Sheets\$pline\$pnumber Rev $revOld.pdf" -Force
#

#Move New Revision from Literature Pending to Literature
Move-item -Path "F:\Documentation Pending\Literature PENDING\$pnumber Rev $revNew.docx" -Destination "F:\Documentation\DOCS\Literature\Product Spec Sheets\$pline\$pnumber Rev $revNew.docx" -Force
#

#Copy New Revision .docx from Literature to Customer Literature
Copy-Item -Path "F:\Documentation\DOCS\Literature\Product Spec Sheets\$pline\$pnumber Rev $revNew.docx" -Destination "F:\Documentation\DOCS\Customer Literature\Spec Sheets\$pline" -Force
#

#Remove Rev Level from new revision
Rename-Item -Path "F:\Documentation\DOCS\Literature\Product Spec Sheets\$pline\$pnumber Rev $revNew.docx" -NewName "F:\Documentation\DOCS\Literature\Product Spec Sheets\$pline\$pnumber.docx" -Force
#

#Convert Customer Literature copy to .pdf, delete .docx file from Customer Literature
ConvertWordTo-PDF -SourceFolder "F:\Documentation\DOCS\Customer Literature\Spec Sheets\$pline"
Remove-Item -Path "F:\Documentation\DOCS\Customer Literature\Spec Sheets\$pline\$pnumber Rev $revNew.docx" -Force
#

#Open Folders for Final Inspection
Invoke-Item -Path "N:\Deep_Archive_Documentation\DOCS\Literature\Product Spec Sheets\$pline\$pnumber"
Invoke-Item -Path "F:\Documentation\DOCS\Literature\Product Spec Sheets\$pline"
#

#endregion


###

#region Revision Notes

#$FileList = Get-ChildItem -Path "F:\Documentation\DOCS\Customer Literature\Spec Sheets\$pline";
#foreach ($File in $FileList) { 
#   $File.Name -match $pnumber;
#    if ($matches) {
#        Remove-Item -Path $File.FullName -Confirm;
#    }
#    $matches = $null
#}
#Old "Delete Old Revision .pdf file from circulation" code 7.14.22

#Rename Old Revision to include Rev Level
#Rename-Item -Path "N:\Deep_Archive_Documentation\DOCS\Literature\Product Spec Sheets\$pline\$pnumber\$pnumber.docx" -NewName "N:\Deep_Archive_Documentation\DOCS\Literature\Product Spec Sheets\$pline\$pnumber\$pnumber Rev $revOld.docx" -WhatIf
#Did not need, Move-Item can change filenames as well.

#endregion