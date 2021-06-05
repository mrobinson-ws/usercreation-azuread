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

#Start While Loop for Quitbox
while ($quitboxOutput -ne "NO"){

#Create Quit Prompt and Close While Loop
$quitboxOutput = [System.Windows.Forms.MessageBox]::Show("Do you need to create another user?" , "User Creation Complete" , 4)
}