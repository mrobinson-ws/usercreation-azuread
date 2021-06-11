#Requires -Modules AzureAD, ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell

#####Declarations#####
#Allow -Verbose to Work
[CmdletBinding()]
Param()
#Include GUI Elements in Script
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()
$quitboxOutput = ""
$LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$SkuToFriendly = @{
    "c42b9cae-ea4f-4ab7-9717-81576235ccac" = "DevPack E5 (No Teams or Audio)"
}
$FriendlyToSku = @{
    "DevPack E5 (No Teams or Audio)" = "c42b9cae-ea4f-4ab7-9717-81576235ccac"
}
#####End of Declarations#####

# Test And Connect To AzureAD If Needed
try {
    Write-Verbose -Message "Testing connection to Azure AD"
    Get-AzureAdDomain -ErrorAction Stop | Out-Null
    Write-Verbose -Message "Already connected to Azure AD"
}
catch {
    Write-Verbose -Message "Connecting to Azure AD"
    Connect-AzureAD
}

#Test And Connect To Microsoft Exchange Online If Needed
try {
    Write-Verbose -Message "Testing connection to Microsoft Exchange Online"
    Get-Mailbox -ErrorAction Stop | Out-Null
    Write-Verbose -Message "Already connected to Microsoft Exchange Online"
}
catch {
    Write-Verbose -Message "Connecting to Microsoft Exchange Online"
    Connect-ExchangeOnline
}



#Start While Loop for Quitbox
while ($quitboxOutput -ne "NO"){
    #####Create User Details Form#####
    $userdetailForm = New-Object System.Windows.Forms.Form
    $userdetailForm.Text = "Please Enter User Details"
    $userdetailForm.Autosize = $true
    $userdetailForm.StartPosition = 'CenterScreen'

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.TabIndex = 5
    $okbutton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $userdetailForm.AcceptButton = $okButton
    $userdetailForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelbutton.TabIndex = 6
    $cancelButton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $userdetailForm.CancelButton = $cancelButton
    $userdetailForm.Controls.Add($cancelButton)

    $passwordTextbox = New-Object System.Windows.Forms.TextBox
    $passwordTextbox.TabIndex = 4
    $passwordTextbox.Dock = [System.Windows.Forms.DockStyle]::Top
    $userdetailForm.Controls.Add($passwordTextbox)
    
    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $passwordLabel.Text = "Password"
    $userdetailForm.Controls.Add($passwordLabel)

    $domainComboBox = New-Object System.Windows.Forms.ComboBox
    $domainComboBox.TabIndex = 3
    $domainComboBox.Dock = [System.Windows.Forms.DockStyle]::Top
    $domainComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach($domain in Get-AzureADDomain){
        $null = $domainComboBox.Items.add($domain.Name)
    }
    $userdetailForm.Controls.Add($domainComboBox)
    
    $domainLabel = New-Object System.Windows.Forms.Label
    $domainLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $domainLabel.Text = "@"
    $userdetailForm.Controls.Add($domainLabel)

    $usernameTextbox = New-Object System.Windows.Forms.TextBox
    $usernameTextbox.TabIndex = 2
    $usernameTextbox.Dock = [System.Windows.Forms.DockStyle]::Top
    $userdetailForm.Controls.Add($usernameTextbox)

    $usernameLabel = New-Object System.Windows.Forms.Label
    $usernameLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $usernameLabel.Text = "Username"
    $userdetailForm.Controls.Add($usernameLabel)

    $lastnameTextbox = New-Object System.Windows.Forms.TextBox
    $lastnameTextbox.TabIndex = 1
    $lastnameTextbox.Dock = [System.Windows.Forms.DockStyle]::Top
    $userdetailForm.Controls.Add($lastnameTextbox)
    
    $lastnameLabel = New-Object System.Windows.Forms.Label
    $lastnameLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $lastnameLabel.Text = "Last Name"
    $userdetailForm.Controls.Add($lastnameLabel)

    $firstnameTextbox = New-Object System.Windows.Forms.TextBox
    $firstnameTextbox.TabIndex = 0
    $firstnameTextbox.Dock = [System.Windows.Forms.DockStyle]::Top
    $userdetailForm.Controls.Add($firstnameTextbox)

    $firstnameLabel = New-Object System.Windows.Forms.Label
    $firstnameLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $firstnameLabel.Text = "First Name"
    $userdetailForm.Controls.Add($firstnameLabel)

    $userdetailForm.Topmost = $true

    $userdetailForm.Add_Shown({$firstnameTextbox.Select()})
    $result = $userdetailForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = $passwordTextbox.text
        $username = $usernameTextbox.Text + "@" + $domainCombobox.Text
        $firstname = $firstnameTextbox.Text
        $lastname = $lastnameTextbox.Text
        $displayname = $firstname + " " + $lastname
    }
    else { Throw }

    $PasswordProfile
    $username
    $firstname
    $lastname
    $displayname

    #####End User Detail Form#####
    
    ##### Create License Selection Form #####
    $Skus = Get-AzureADSubscribedSku | Select-Object -Property Sku*,ConsumedUnits -ExpandProperty PrepaidUnits

    $LicenseSelectWindow = New-Object System.Windows.Forms.Form
    $LicenseSelectWindow.Text = "Select Licenses"
    $LicenseSelectWindow.AutoSize = $true
    $LicenseSelectWindow.AutoSizeMode = "GrowAndShrink"
    $LicenseSelectWindow.MinimizeBox = $false
    $LicenseSelectWindow.MaximizeBox = $false
    $LicenseSelectWindow.StartPosition = "CenterScreen"
    $LicenseSelectWindow.FormBorderStyle = "Fixed3D"

    $CheckedListBox = New-Object System.Windows.Forms.CheckedListBox
    $CheckedListBox.AutoSize = $true
    $CheckedListBox.CheckOnClick = $true #so we only have to click once to check a box
    foreach ($Sku in $Skus) {
        Clear-Variable HRSku -ErrorAction SilentlyContinue
        $HRSku = $SkuToFriendly.Item("$($Sku.SkuID)")
        $CheckedListBoxOutput = $HRSku + " -- " + ($Sku.Enabled-$License.ConsumedUnits) + " of " + $Sku.Enabled + " Available"
        $null = $CheckedListBox.Items.Add($CheckedListBoxOutput)
    }
    $CheckedListBox.ClearSelected()
    $CheckedListBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $LicenseSelectWindow.Controls.Add($CheckedListBox)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Text = "Use Selected Licenses"
    $OKButton.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $LicenseSelectWindow.Controls.Add($OKButton)

    #put form in front of other windows
    $LicenseSelectWindow.TopMost = $true

    #display the form
    $null = $LicenseSelectWindow.ShowDialog()
    if ($OKButton.DialogResult -eq "OK") {
        $CheckedListBox.CheckedItems
    }

    foreach($checkedlicense in $CheckedListBox.CheckedItems){
        $converttosku = $checkedlicense -replace '\s--\s.*'
        $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $License.SkuID = $FriendlyToSku.Item("$($converttosku)")
        $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $LicensesToAssign.AddLicenses = $License
        #Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $LicensesToAssign
    }
#Create Quit Prompt and Close While Loop
$quitboxOutput = [System.Windows.Forms.MessageBox]::Show("Do you need to create another user?" , "User Creation Complete" , 4)
}