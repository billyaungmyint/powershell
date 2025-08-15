# PowerShell script to get user and mailbox details from Microsoft Entra and Exchange Online
# Version 1.0
#
# This script retrieves information for a specified user, including:
# - General user details from Microsoft Entra (Name, Email, Department, Account Status, Creation Date)
# - Mailbox details from Exchange Online (Mailbox Size, Archive Size, Recoverable Items Size, Retention Policy)
# - Audit information (User creation and deletion events)
#
# Requirements:
# - PowerShell 5.1 or later
# - Microsoft.Graph module
# - ExchangeOnlineManagement module
# - Permissions to read user information from Microsoft Entra and Exchange Online.
# - For audit log information, 'AuditLog.Read.All' permission is required for Microsoft Graph.

# --- Functions ---

# Function to check and install required modules
function Ensure-Modules {
    $requiredModules = @("Microsoft.Graph", "ExchangeOnlineManagement")

    foreach ($module in $requiredModules) {
        if (Get-Module -ListAvailable -Name $module) {
            Write-Host "$module module is already installed."
        } else {
            Write-Host "$module module is not installed. Attempting to install..."
            try {
                Install-Module -Name $module -Repository PSGallery -Force -AllowClobber -Scope CurrentUser
                Write-Host "$module module has been successfully installed."
            }
            catch {
                Write-Error "Error installing $module. Please install it manually and re-run the script."
                throw
            }
        }
    }
}

# Function to connect to services
function Connect-Services {
    try {
        # Connect to Microsoft Graph
        Write-Host "Connecting to Microsoft Graph..."
        Connect-MgGraph -Scopes "User.Read.All, AuditLog.Read.All"
        Write-Host "Successfully connected to Microsoft Graph."

        # Connect to Exchange Online
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "Successfully connected to Exchange Online."
    }
    catch {
        Write-Error "Failed to connect to one or more services. $_"
        throw
    }
}

# Function to get user details
function Get-UserDetails {
    param (
        [string]$UserPrincipalName
    )

    try {
        Write-Host "Getting user details for $UserPrincipalName from Microsoft Entra..."
        $userDetails = Get-MgUser -UserId $UserPrincipalName -Property "DisplayName, UserPrincipalName, Department, AccountEnabled, CreatedDateTime"

        if ($userDetails) {
            $output = [PSCustomObject]@{
                "Name"            = $userDetails.DisplayName
                "Email"           = $userDetails.UserPrincipalName
                "Department"      = $userDetails.Department
                "Account Status"  = if ($userDetails.AccountEnabled) { "Enabled" } else { "Disabled" }
                "Created Date"    = $userDetails.CreatedDateTime
            }

            # Get mailbox details from Exchange Online
            Write-Host "Getting mailbox details for $UserPrincipalName from Exchange Online..."
            $mailbox = Get-EXOMailbox -Identity $UserPrincipalName
            $mailboxStats = Get-EXOMailboxStatistics -Identity $UserPrincipalName

            $output | Add-Member -MemberType NoteProperty -Name "Mailbox Size" -Value $mailboxStats.TotalItemSize
            $output | Add-Member -MemberType NoteProperty -Name "Archive Mailbox Size" -Value $mailboxStats.TotalArchiveSize
            $output | Add-Member -MemberType NoteProperty -Name "Recoverable Items Size" -Value $mailboxStats.TotalDeletedItemSize
            $output | Add-Member -MemberType NoteProperty -Name "Mailbox Retention Policy" -Value $mailbox.RetentionPolicy

            # Get Audit Information
            Write-Host "Querying audit logs for creation and deletion events..."
            $auditInfo = Get-AuditInfo -UserPrincipalName $UserPrincipalName
            if ($auditInfo) {
                $output | Add-Member -MemberType NoteProperty -Name "Deleted Date" -Value $auditInfo.DeletedDate
                $output | Add-Member -MemberType NoteProperty -Name "Deleted By" -Value $auditInfo.DeletedBy
            }

            # Display the collected information
            $output | Format-List
        } else {
            Write-Warning "User with UPN $UserPrincipalName not found."
        }
    }
    catch {
        Write-Error "An error occurred while retrieving user details: $_"
    }
}

# Function to get audit log information
function Get-AuditInfo {
    param (
        [string]$UserPrincipalName
    )

    # Note: This requires AuditLog.Read.All permission for Microsoft Graph.
    # It can also take a long time to query audit logs.

    $auditOutput = @{
        "DeletedDate" = "N/A"
        "DeletedBy"   = "N/A"
    }

    # Query for user deletion event
    $deleteFilter = "activityDisplayName eq 'Delete user' and result eq 'success' and targetResources/any(c:c/userPrincipalName eq '$UserPrincipalName')"
    $deleteLog = Get-MgAuditLogDirectoryAudit -Filter $deleteFilter

    if ($deleteLog) {
        $auditOutput.DeletedDate = $deleteLog[0].ActivityDateTime
        $auditOutput.DeletedBy = ($deleteLog[0].InitiatedBy.User.UserPrincipalName | Select-Object -First 1)
    }

    return $auditOutput
}


# Function to disconnect from services
function Disconnect-Services {
    Write-Host "Disconnecting from all services..."
    if (Get-MgConnection) {
        Disconnect-MgGraph
        Write-Host "Disconnected from Microsoft Graph."
    }
    $exchangeSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' }
    if ($exchangeSession) {
        Remove-PSSession $exchangeSession
        Write-Host "Disconnected from Exchange Online."
    }
}


# --- Main Script ---

# Main execution block
function Start-Script {
    try {
        Write-Host "Starting script..."

        # Step 1: Ensure modules are installed
        Ensure-Modules

        # Step 2: Connect to services
        Connect-Services

        # Step 3: Get user input
        $upn = Read-Host "Please enter the User Principal Name (UPN) of the user"

        # Step 4: Get and display user details
        Get-UserDetails -UserPrincipalName $upn

    }
    catch {
        Write-Error "An error occurred: $_"
    }
    finally {
        # Step 5: Disconnect from services
        Disconnect-Services
        Write-Host "Script finished."
    }
}

# --- Script Entry Point ---
Start-Script
