# PowerShell script to get user and mailbox details using Azure AD and Exchange Online modules
# Version 2.0 - No Microsoft Graph dependency
#
# This script retrieves information for a specified user, including:
# - General user details from Azure AD (Name, Email, Department, Account Status, Creation Date)
# - Mailbox details from Exchange Online (Mailbox Size, Archive Size, Recoverable Items Size, Retention Policy)
# - Basic audit information from available cmdlets
#
# Requirements:
# - PowerShell 5.1 or later
# - AzureAD module (legacy but widely supported)
# - ExchangeOnlineManagement module
# - Basic user read permissions

# --- Functions ---

# Function to check and install required modules
function Ensure-Modules {
    $requiredModules = @("AzureAD", "ExchangeOnlineManagement")

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
        # Connect to Azure AD
        Write-Host "Connecting to Azure AD..."
        Connect-AzureAD
        Write-Host "Successfully connected to Azure AD."

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
        Write-Host "Getting user details for $UserPrincipalName from Azure AD..."
        $userDetails = Get-AzureADUser -ObjectId $UserPrincipalName

        if ($userDetails) {
            $output = [PSCustomObject]@{
                "Name"            = $userDetails.DisplayName
                "Email"           = $userDetails.UserPrincipalName
                "Department"      = $userDetails.Department
                "Job Title"       = $userDetails.JobTitle
                "Account Status"  = if ($userDetails.AccountEnabled) { "Enabled" } else { "Disabled" }
                "Created Date"    = $userDetails.ExtensionProperty.createdDateTime
                "Last Sign-In"    = $userDetails.RefreshTokensValidFromDateTime
            }

            # Get mailbox details from Exchange Online
            Write-Host "Getting mailbox details for $UserPrincipalName from Exchange Online..."
            try {
                $mailbox = Get-EXOMailbox -Identity $UserPrincipalName -ErrorAction Stop
                $mailboxStats = Get-EXOMailboxStatistics -Identity $UserPrincipalName -ErrorAction Stop

                $output | Add-Member -MemberType NoteProperty -Name "Mailbox Size" -Value $mailboxStats.TotalItemSize
                $output | Add-Member -MemberType NoteProperty -Name "Item Count" -Value $mailboxStats.ItemCount
                $output | Add-Member -MemberType NoteProperty -Name "Archive Mailbox Size" -Value $mailboxStats.TotalArchiveSize
                $output | Add-Member -MemberType NoteProperty -Name "Recoverable Items Size" -Value $mailboxStats.TotalDeletedItemSize
                $output | Add-Member -MemberType NoteProperty -Name "Mailbox Retention Policy" -Value $mailbox.RetentionPolicy
                $output | Add-Member -MemberType NoteProperty -Name "Mailbox Database" -Value $mailbox.Database
                $output | Add-Member -MemberType NoteProperty -Name "Mailbox Type" -Value $mailbox.RecipientTypeDetails
            }
            catch {
                Write-Warning "Could not retrieve mailbox information: $_"
                $output | Add-Member -MemberType NoteProperty -Name "Mailbox Status" -Value "No mailbox found or access denied"
            }

            # Get additional Exchange recipient information
            Write-Host "Getting additional recipient information..."
            try {
                $recipient = Get-EXORecipient -Identity $UserPrincipalName -ErrorAction SilentlyContinue
                if ($recipient) {
                    $output | Add-Member -MemberType NoteProperty -Name "Primary SMTP Address" -Value $recipient.PrimarySmtpAddress
                    $output | Add-Member -MemberType NoteProperty -Name "Email Addresses" -Value ($recipient.EmailAddresses -join "; ")
                }
            }
            catch {
                Write-Warning "Could not retrieve recipient information: $_"
            }

            # Get basic audit information from available sources
            Write-Host "Getting available audit information..."
            $auditInfo = Get-BasicAuditInfo -UserPrincipalName $UserPrincipalName
            if ($auditInfo.LastLogon -ne "N/A") {
                $output | Add-Member -MemberType NoteProperty -Name "Last Mailbox Logon" -Value $auditInfo.LastLogon
            }

            # Display the collected information
            Write-Host "`n=== USER AND MAILBOX DETAILS ===" -ForegroundColor Green
            $output | Format-List
        } else {
            Write-Warning "User with UPN $UserPrincipalName not found."
        }
    }
    catch {
        Write-Error "An error occurred while retrieving user details: $_"
    }
}

# Function to get basic audit information from available sources
function Get-BasicAuditInfo {
    param (
        [string]$UserPrincipalName
    )

    $auditOutput = @{
        "LastLogon" = "N/A"
    }

    try {
        # Try to get last logon information from mailbox statistics
        $mailboxStats = Get-EXOMailboxStatistics -Identity $UserPrincipalName -ErrorAction SilentlyContinue
        if ($mailboxStats -and $mailboxStats.LastLogonTime) {
            $auditOutput.LastLogon = $mailboxStats.LastLogonTime
        }
    }
    catch {
        Write-Verbose "Could not retrieve mailbox statistics for audit info: $_"
    }

    return $auditOutput
}

# Function to disconnect from services
function Disconnect-Services {
    Write-Host "Disconnecting from all services..."
    
    try {
        if (Get-AzureADCurrentSessionInfo -ErrorAction SilentlyContinue) {
            Disconnect-AzureAD
            Write-Host "Disconnected from Azure AD."
        }
    }
    catch {
        Write-Verbose "Azure AD was not connected or already disconnected."
    }

    try {
        $exchangeSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' }
        if ($exchangeSession) {
            Remove-PSSession $exchangeSession
            Write-Host "Disconnected from Exchange Online."
        }
    }
    catch {
        Write-Verbose "Exchange Online was not connected or already disconnected."
    }
}

# Function to display help information
function Show-Help {
    Write-Host @"

=== PowerShell User and Mailbox Details Script (No Graph) ===

This script retrieves comprehensive user and mailbox information using:
- Azure AD PowerShell module (legacy but widely supported)
- Exchange Online Management module

USAGE:
    Start-Script                    # Interactive mode - prompts for UPN
    Get-UserDetails -UserPrincipalName "user@domain.com"  # Direct function call

REQUIREMENTS:
- PowerShell 5.1 or later
- AzureAD module
- ExchangeOnlineManagement module
- Basic user read permissions (no admin consent required)

INFORMATION RETRIEVED:
- User details (name, email, department, job title, account status)
- Mailbox size and statistics
- Retention policies
- Email addresses
- Last logon information (when available)

"@ -ForegroundColor Cyan
}

# --- Main Script ---

# Main execution block
function Start-Script {
    param (
        [string]$UserPrincipalName,
        [switch]$Help
    )

    if ($Help) {
        Show-Help
        return
    }

    try {
        Write-Host "=== Starting User and Mailbox Details Script ===" -ForegroundColor Green

        # Step 1: Ensure modules are installed
        Write-Host "`nStep 1: Checking required modules..." -ForegroundColor Yellow
        Ensure-Modules

        # Step 2: Connect to services
        Write-Host "`nStep 2: Connecting to services..." -ForegroundColor Yellow
        Connect-Services

        # Step 3: Get user input if not provided
        if (-not $UserPrincipalName) {
            Write-Host "`nStep 3: Getting user input..." -ForegroundColor Yellow
            $UserPrincipalName = Read-Host "Please enter the User Principal Name (UPN) of the user"
        }

        # Step 4: Validate input
        if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
            throw "User Principal Name cannot be empty."
        }

        # Step 5: Get and display user details
        Write-Host "`nStep 4: Retrieving user and mailbox details..." -ForegroundColor Yellow
        Get-UserDetails -UserPrincipalName $UserPrincipalName

    }
    catch {
        Write-Error "An error occurred: $_"
        Write-Host "Use 'Start-Script -Help' for usage information." -ForegroundColor Yellow
    }
    finally {
        # Step 6: Disconnect from services
        Write-Host "`nStep 5: Cleaning up connections..." -ForegroundColor Yellow
        Disconnect-Services
        Write-Host "`n=== Script finished ===" -ForegroundColor Green
    }
}

# --- Script Entry Point ---

# Check if script is being run directly or dot-sourced
if ($MyInvocation.InvocationName -ne '.') {
    # Script is being run directly
    Start-Script
} else {
    # Script is being dot-sourced, show help
    Write-Host "Script loaded. Available functions:" -ForegroundColor Green
    Write-Host "  Start-Script [-UserPrincipalName <UPN>] [-Help]" -ForegroundColor Cyan
    Write-Host "  Get-UserDetails -UserPrincipalName <UPN>" -ForegroundColor Cyan
    Write-Host "  Show-Help" -ForegroundColor Cyan
}