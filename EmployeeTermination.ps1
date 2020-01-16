 $wshell = New-Object -ComObject Wscript.Shell
 $wshell.Popup("Please enter in your Domain Admin credentials.  **removed**.",0,"Credentials Needed!",0x0)	
 $creds = Get-Credential
 $PSDefaultParameterValues = @{"*-AD*:Credential"=$creds}

#Create the connection to the exchange server
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://**removed**/ #-Authentication Kerberos
Import-PSSession $ExchangeSession

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 
	
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Terminated Employee Process Form"
$objForm.Size = New-Object System.Drawing.Size(500,400) 
$objForm.StartPosition = "CenterScreen"
$objForm.MaximizeBox = $False


$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$userinput=$UserTextBox.Text;$forwardemail=$ForwardingTextBox.Text;$ticketnumber=$TicketTextBox.Text;$disableuser=$DisableUserCheckbox.Checked;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$Font = New-Object System.Drawing.Font("Verdana",8,[System.Drawing.FontStyle]::Bold) 
#$objForm.Font = $Font 
#VERSION NUMBER
$VersionLabel = New-Object System.Windows.Forms.Label
$VersionLabel.Location = New-Object System.Drawing.Size(450,10) 
$VersionLabel.Size = New-Object System.Drawing.Size(120,20) 
$VersionLabel.Font = $Font 
$VersionLabel.Text = "V2"
$objForm.Controls.Add($VersionLabel) 

#OK AND CANCEL BUTTONS
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(75,320)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Add_Click({$userinput=$UserTextBox.Text;$ticketnumber=$TicketTextBox.Text;$forwardemail=$ForwardingTextBox.Text;$disableuser=$DisableUserCheckbox.Checked;$objForm.Close()})
$objForm.Controls.Add($OKButton)


#USERNAME LABEL
$UserLabel = New-Object System.Windows.Forms.Label
$UserLabel.Location = New-Object System.Drawing.Size(10,20) 
$UserLabel.Size = New-Object System.Drawing.Size(280,20) 
$UserLabel.Text = "Username of Employee to be terminated"
$objForm.Controls.Add($UserLabel) 
#USERNAME TEXT BOX
$UserTextBox = New-Object System.Windows.Forms.TextBox 
$UserTextBox.Location = New-Object System.Drawing.Size(10,40) 
$UserTextBox.Size = New-Object System.Drawing.Size(180,20) 
$objForm.Controls.Add($UserTextBox) 

#DISABLE USER CHECKBOX CONTROL
#$DisableUserCheckbox = New-Object System.Windows.Forms.Checkbox 
#$DisableUserCheckbox.Location = New-Object System.Drawing.Size(220,30) 
#$DisableUserCheckbox.Size = New-Object System.Drawing.Size(140,40)
#$DisableUserCheckbox.Text = "Disable The User?"
#$objForm.Controls.Add($DisableUserCheckbox)

#EXTERNAL USER CHECKBOX CONTROL
$ExternalUserCheckbox = New-Object System.Windows.Forms.Checkbox 
$ExternalUserCheckbox.Location = New-Object System.Drawing.Size(120,70) 
$ExternalUserCheckbox.Size = New-Object System.Drawing.Size(120,40)
$ExternalUserCheckbox.Text = "External User?"
$objForm.Controls.Add($ExternalUserCheckbox)

<#
#FORWARD EMAIL LABEL
$FowardEmailLabel = New-Object System.Windows.Forms.Label
$FowardEmailLabel.Location = New-Object System.Drawing.Size(10,80) 
$FowardEmailLabel.Size = New-Object System.Drawing.Size(280,20)
$FowardEmailLabel.Text = "Forward Email to Manager? If Yes, Type In Email Address"
$objForm.Controls.Add($FowardEmailLabel)

#FORWARD EMAIL TEXT BOX
$ForwardingTextBox = New-Object System.Windows.Forms.TextBox 
$ForwardingTextBox.Location = New-Object System.Drawing.Size(10,100) 
$ForwardingTextBox.Size = New-Object System.Drawing.Size(180,40) 
$objForm.Controls.Add($ForwardingTextBox) 

#ENTER TICKET NUMBER TEXT LABEL
$TicketLabel = New-Object System.Windows.Forms.Label
$TicketLabel.Location = New-Object System.Drawing.Size(10,150) 
$TicketLabel.Size = New-Object System.Drawing.Size(80,20)
$TicketLabel.Text = "Issue Number"
$objForm.Controls.Add($TicketLabel)

$TicketTextBox = New-Object System.Windows.Forms.TextBox 
$TicketTextBox.Location = New-Object System.Drawing.Size(10,170) 
$TicketTextBox.Size = New-Object System.Drawing.Size(40,250) 
$objForm.Controls.Add($TicketTextBox) 
#>

#CANCEL BUTTONS
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(350,320)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close(); $cancel = $true})
$objForm.Controls.Add($CancelButton)


$objForm.Topmost = $True
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()
#if ($cancel) {return}
#$OKButton.Add_Click({$userinput=$UserTextBox.Text;$ticketnumber=$TicketTextBox.Text;$forwardemail=$ForwardingTextBox.Text;$disableuser=$DisableUserCheckbox.Checked;$objForm.Close()})
#$CancelButton.Add_Click({$objForm.Close()})

#COMMON GLOBAL VARIABLES
$ExternalUserCheckbox=$ExternalUserCheckbox.Checked
$disableusercheckbox=$DisableUserCheckbox.Checked
$userinput=$UserTextBox.Text
#$forwardemail=$ForwardingTextBox.Text
#$ticketnumber=$TicketTextBox.Text

$Month = Get-Date -format MM
$Day = Get-Date -format dd
$Year = Get-Date -format yyyy


#######-------------------------------------------------------------------------
try {
    If ($cancel) {

    Write-host "Cancelled" -ForegroundColor Green
    return

    }

    If ($OKButton.Add_Click) {

    
    ########
    #ACTIVE DIRECTORY ACTIONS
    #########

    <#DISABLE THE USER
    If ($disableusercheckbox -eq $true)
    {
      Disable-ADAccount -Identity $userinput
      Write-host "$userinput has been disabled" -ForegroundColor Green
    } else { 
	    Write-host "$userinput has not been disabled at this time" 
    }
    #>

    #DISABLE THE USER
    Disable-ADAccount -Identity $userinput
    Write-host "$userinput has been disabled" -ForegroundColor Green

    #GETS ALL GROUPS USER WAS PART OF BEFORE BLOWING THEM OUT
        $User = $userinput
        $List=@()
        $Groups = Get-ADUser -Identity $User -Properties * | select -ExpandProperty memberof
        foreach($i in $Groups){
        $i = ($i -split ',')[0]
        $List += "`r`n" + ($i -creplace 'CN=|}','')
        }
    
    #BLOW OUT GROUPS OF USER EXCEPT DOMAIN USERS
    If ($ExternalUserCheckbox -eq $true)
    {
        write-host -NoNewline "$userinput is an " -ForegroundColor Green; Write-Host "External account" -ForegroundColor Yellow
        Add-adgroupmember -Identity "Domain Users" -members $userinput -Confirm:$False #ADD USER TO THE DOMAIN USERS GROUP
        $Group = get-adgroup "CN=Domain Users,CN=Users,DC=ltcpartners,DC=com" -properties @("primaryGroupToken") #GET DOMAIN USERS PRIMARY TOKEN
        write-host "Setting Domain Users OU as the primary group.." -ForegroundColor Green
        Set-ADUser -Identity $userinput -Replace @{primarygroupid=$group.primaryGroupToken} #SET DOMAIN USERS TOKEN TO THE USER
        (get-aduser $userinput -properties memberof).memberof|remove-adgroupmember -member $userinput -Confirm:$False #REMOvE FROM ALL GROUPS EXCEPT DOMAIN
        write-host "Removing all groups.." -ForegroundColor Green
    } else { 
        write-host -NoNewline "$userinput is a " -ForegroundColor Green; Write-Host "Domain account" -ForegroundColor Yellow
        (get-aduser $userinput -properties memberof).memberof|remove-adgroupmember -member $userinput -Confirm:$False
        Write-Host "Removing all groups.." -ForegroundColor Green 
    }


    #SETS THE USERS TITLE,COMPANY/MANAGER TO DISABLED
    set-aduser -identity $userinput -description "LTCP - Disabled at $Month/$Day/$Year"
    set-aduser -identity $userinput -company $null
    set-aduser -identity $userinput -manager $null
    set-aduser -identity $userinput -department $null
    #set-aduser -identity $userinput -description "LTCP - Disabled at $Month/$Day/$Year"
    Write-Host "Deleting Company, Manager, and Department field." -ForegroundColor Green

    #CHANGES THE USERS PASSWORD
    $newpwd = ConvertTo-SecureString -String "G00dBye@1234" -AsPlainText –Force
    Set-ADAccountPassword $userinput –NewPassword $newpwd -Reset

    #MOVES THE USER TO DISABLED USERS
    If ($ExternalUserCheckbox -eq $true)
    {
        Get-ADUser -Filter { samAccountName -like $userinput } | Move-ADObject –TargetPath "OU=Disabled External Users,OU=Disabled User Accounts,DC=ltcpartners,DC=com"
        #$disabled = $userinput + " has been moved to disabled external group"
        write-host -NoNewline "$userinput has been moved to " -ForegroundColor Green; Write-Host "External User group" -ForegroundColor Yellow
    } else { 
        Get-ADUser -Filter { samAccountName -like $userinput } | Move-ADObject –TargetPath "OU=Disabled User Accounts,DC=ltcpartners,DC=com"
        #$notdisabled = $userinput + " has not been disabled at this time" 
        Write-Host -NoNewline "$userinput has been moved to " -ForegroundColor Green; Write-Host "Disabled User group" -ForegroundColor Yellow
    }

    #HIDES USER FROM GLOBAL ADDRESS BOOK and configures forwarding
    #Set-Mailbox -Identity $userinput -ForwardingAddress $forwardemail -HiddenFromAddressListsEnabled $true

    }
} catch {
     Write-host "`"Uh oh something went wrong..`" -Definitely not Efras" -ForegroundColor Red
    
} finally {
    #REMOVES THE SESSION
    Remove-PSsession $ExchangeSession 

    Start-Sleep -s 10
}
