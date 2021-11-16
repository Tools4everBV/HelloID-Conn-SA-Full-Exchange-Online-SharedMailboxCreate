$path = "OU=Shared,OU=Mailbox Permissies,OU=Groepen,OU=enyoi,DC=org"

# Fixed values
$AutoMapping = $false

# Connect to Office 365
try{
    Hid-Write-Status -Event Information -Message "Connecting to Office 365.."

    $module = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential ($ExchangeOnlineAdminUsername, $securePassword)

    $exchangeSession = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -TrackPerformance:$false -ErrorAction Stop 

    Hid-Write-Status -Event Success -Message "Successfully connected to Office 365"
}catch{
    throw "Could not connect to Exchange Online, error: $_"
}

try{
    # Create Shared mailbox
    try{
        $SharedMailboxParams = @{
            Name                =   $Name
            DisplayName         =   $Name
            PrimarySmtpAddress  =   $Alias.Replace(" ","") + "@$Maildomain"
            Alias               =   $Alias.Replace(" ","")
        }

        $SharedMailbox = New-Mailbox -Shared @SharedMailboxParams -ErrorAction Stop
     
        Hid-Write-Status -Message "Shared Mailbox [$($SharedMailboxParams.Name)] created successfully" -Event Success
        HID-Write-Summary -Message "Shared Mailbox [$($SharedMailboxParams.Name)] created successfully" -Event Success
    } catch {
        HID-Write-Status -Message "Error creating Shared Mailbox [$($SharedMailboxParams.Name)]. Error: $($_)" -Event Error
        HID-Write-Summary -Message "Error creating Shared Mailbox [$($SharedMailboxParams.Name)]" -Event Failed
    }

    # Create AD Group
    try{   
        $groupParams = @{
            Name            = "MBX_" + $Name.Replace(' ','_')
            GroupScope      = "Universal"
            Path            = $path
        }
 
        $primarySmtpAddress = $groupParams.name.Replace(" ","") + "@$Maildomain"
        $adGroup = New-ADGroup @groupParams -OtherAttributes @{'mail'=$primarySmtpAddress} -ErrorAction Stop

        Hid-Write-Status -Message "AD Group [$($groupParams.Name)] created successfully" -Event Success
        HID-Write-Summary -Message "AD Group [$($groupParams.Name)] created successfully" -Event Success
    } catch {
        HID-Write-Status -Message "Error creating AD Group [$($groupParams.Name)]. Error: $($_)" -Event Error
        HID-Write-Summary -Message "Error creating AD Group [$($groupParams.Name)]" -Event Failed
    }


    try{
        Hid-Write-Status -Message "Checking if mailbox with name '$($SharedMailboxParams.Name)' exists..." -Event Warning
        $mailboxCheck = $null
        do {
            try{
                $mailboxCheck = Get-Mailbox $SharedMailboxParams.Name -ErrorAction Stop
                $identity = $mailboxCheck.Identity
            }catch{
                Start-Sleep -Seconds 30
            }
        }while ($mailboxCheck -eq $null)    


        Hid-Write-Status -Message "Checking if AD group with name '$($groupParams.Name)' exists..." -Event Warning
        $adGroupCheck = $null
        do {
            try{
                $adGroupCheck = Get-Group -Identity $groupParams.Name -ErrorAction Stop
                $group = $adGroupCheck.Name
            }catch{
                Start-Sleep -Seconds 30
            }
        }while ($adGroupCheck -eq $null)


        # Add Full Access Permissions
        if($AutoMapping){
            $mailboxPermission = Add-MailboxPermission -Identity $identity -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -User $group -ErrorAction Stop
        }else{
            $mailboxPermission = Add-MailboxPermission -Identity $identity -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -User $group -ErrorAction Stop
        }

        Hid-Write-Status -Message "Assigned access rights [FullAccess] for mailbox [$($identity)] to [$($group)] successfully" -Event Success
        HID-Write-Summary -Message "Assigned access rights [FullAccess] for mailbox [$($identity)] to [$($group)] successfully" -Event Success

        # Add Send As Permissions
        $recipientPermission = Add-RecipientPermission -Identity $identity -Trustee $group -AccessRights SendAs -Confirm:$false

        Hid-Write-Status -Message "Assigned access rights [SendAs] for mailbox [$($identity)] to [$($group)] successfully" -Event Success
        HID-Write-Summary -Message "Assigned access rights [SendAs] for mailbox [$($identity)] to [$($group)] successfully" -Event Success
    } catch {
        HID-Write-Status -Message "Error assigning access rights [FullAccess] and [SendAs] for mailbox [$($identity)] to [$($group)]. Error: $($_)" -Event Error
        HID-Write-Summary -Message "Error assigning access rights [FullAccess] and [SendAs] for mailbox [$($identity)] to [$($group)]" -Event Failed
    }
} finally {
    Hid-Write-Status -Event Information -Message "Disconnecting from Office 365.."
    $exchangeSessionEnd = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false -ErrorAction Stop
    Hid-Write-Status -Event Success -Message "Successfully disconnected from Office 365"
}
