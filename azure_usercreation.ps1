#Requires -Modules AzureAD, ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell

#####Declarations#####
#Allow -Verbose to Work
[CmdletBinding()]
Param()
#Include GUI Elements in Script
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Windows.Forms.Application]::EnableVisualStyles()
#Set variable for loop to function
$quitboxOutput = ""
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

$SkuArray = [System.Collections.ArrayList]@()

#Start While Loop for Quitbox
while ($quitboxOutput -ne "NO"){
    ##### Create License Selection Form #####
    $licenses = Get-AzureADSubscribedSku | Select-Object -Property Sku*,ConsumedUnits -ExpandProperty PrepaidUnits

    $LicenseSelectWindow = New-Object System.Windows.Forms.Form
    $LicenseSelectWindow.Text = "Select Licenses"
    $LicenseSelectWindow.AutoSize = $true
    $LicenseSelectWindow.AutoSizeMode = "GrowAndShrink"
    $LicenseSelectWindow.MinimizeBox = $false
    $LicenseSelectWindow.MaximizeBox = $false
    $LicenseSelectWindow.StartPosition = "CenterScreen"
    $LicenseSelectWindow.FormBorderStyle = 'Fixed3D'

    $CheckedListBox = New-Object System.Windows.Forms.CheckedListBox
    $CheckedListBox.AutoSize = $true
    $CheckedListBox.CheckOnClick = $true #so we only have to click once to check a box
    foreach ($license in $licenses) {
        $CheckedListBoxOutput = $license.SkuPartNumber + " -- " + ($license.Enabled-$License.ConsumedUnits) + " of " + $license.Enabled + " Available"
        $CheckedListBox.Items.Add($CheckedListBoxOutput)
        $SkuArray.Add($license.SkuPartNumber)
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
    $DisplayForm = $LicenseSelectWindow.ShowDialog()
    if ($OKButton.DialogResult -eq "OK") {
        $SkuArray
    }

#Create Quit Prompt and Close While Loop
$quitboxOutput = [System.Windows.Forms.MessageBox]::Show("Do you need to create another user?" , "User Creation Complete" , 4)
}