<#

.SYNOPSIS

    This function creates a new user with the same licenses and group memberships as an existing user.

.DESCRIPTION

    This function creates a new user in the tenant with the same licenses and group memberships as an existing user.

    The function requires the following environment variables to be set:

    Ms365_AuthAppId - Application Id of the service principal
    Ms365_AuthSecretId - Secret Id of the service principal
    Ms365_TenantId - Tenant Id of the Microsoft 365 tenant
    SecurityKey - Optional, use this as an additional step to secure the function

    The function requires the following modules to be installed:

    Microsoft.Graph

.INPUTS

    NewUserEmail - new user email address
    ExistingUserEmail - existing user email address to copy from
    NewUserFirstName - first name of new user
    NewUserLastName - last name of new user
    NewUserDisplayName - display name of new user
    TenantId - string value of the tenant id, if blank uses the environment variable Ms365_TenantId
    TicketId - optional - string value of the ticket id used for transaction tracking
    SecurityKey - Optional, use this as an additional step to secure the function

    JSON Structure

    {
        "NewUserEmail": "newuser@domain.com",
        "ExistingUserEmail": "existinguser@domain.com",
        "NewUserFirstName": "John",
        "NewUserLastName": "Doe",
        "NewUserDisplayName": "John Doe",
        "TenantId": "12345678-1234-1234-123456789012",
        "TicketId": "123456",
        "SecurityKey": "optional"
    }

.OUTPUTS

    JSON response with the following fields:

    Message - Descriptive string of result
    TicketId - TicketId passed in Parameters
    ResultCode - 200 for success, 500 for failure
    ResultStatus - "Success" or "Failure"

#>

using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "Create User Like Another User function triggered."

$resultCode = 200
$message = ""

$NewUserEmail = $Request.Body.NewUserEmail
$ExistingUserEmail = $Request.Body.ExistingUserEmail
$NewUserFirstName = $Request.Body.NewUserFirstName
$NewUserLastName = $Request.Body.NewUserLastName
$NewUserDisplayName = $Request.Body.NewUserDisplayName
$TenantId = $Request.Body.TenantId
$TicketId = $Request.Body.TicketId
$SecurityKey = $env:SecurityKey

if ($SecurityKey -And $SecurityKey -ne $Request.Headers.SecurityKey) {
    Write-Host "Invalid security key"
    break;
}

if (-Not $NewUserEmail) {
    $message = "NewUserEmail cannot be blank."
    $resultCode = 500
}
else {
    $NewUserEmail = $NewUserEmail.Trim()
}

if (-Not $ExistingUserEmail) {
    $message = "ExistingUserEmail cannot be blank."
    $resultCode = 500
}
else {
    $ExistingUserEmail = $ExistingUserEmail.Trim()
}

if (-Not $NewUserFirstName) {
    $message = "NewUserFirstName cannot be blank."
    $resultCode = 500
}

if (-Not $NewUserLastName) {
    $message = "NewUserLastName cannot be blank."
    $resultCode = 500
}

if (-Not $NewUserDisplayName) {
    $NewUserDisplayName = "$NewUserFirstName $NewUserLastName"
}

if (-Not $TenantId) {
    $TenantId = $env:Ms365_TenantId
}
else {
    $TenantId = $TenantId.Trim()
}

if (-Not $TicketId) {
    $TicketId = ""
}

Write-Host "New User Email: $NewUserEmail"
Write-Host "Existing User Email: $ExistingUserEmail"
Write-Host "New User First Name: $NewUserFirstName"
Write-Host "New User Last Name: $NewUserLastName"
Write-Host "New User Display Name: $NewUserDisplayName"
Write-Host "Tenant Id: $TenantId"
Write-Host "Ticket Id: $TicketId"

if ($resultCode -Eq 200) {
    $secure365Password = ConvertTo-SecureString -String $env:Ms365_AuthSecretId -AsPlainText -Force
    $credential365 = New-Object System.Management.Automation.PSCredential($env:Ms365_AuthAppId, $secure365Password)

    Connect-MgGraph -ClientSecretCredential $credential365 -TenantId $TenantId

    # Define the existing user's UserPrincipalName (UPN) and the new user's UPN
    $existingUserUpn = $ExistingUserEmail
    $newUserUpn = $NewUserEmail

    # Retrieve the existing user's details including licenses
    $existingUser = Get-MgUser -UserId $existingUserUpn -Property Id,DisplayName,UserPrincipalName,AssignedLicenses

    if (-Not $existingUser) {
        $message = "Request failed. User `"$ExistingUserEmail`" could not be found."
        $resultCode = 500
    }
    else {
        # Check if the existing user has any assigned licenses
        if ($existingUser.AssignedLicenses.Count -eq 0) {
            Write-Host "The existing user `"$ExistingUserEmail`" does not have any assigned licenses."
        }
        else {
            Write-Host "Existing user has $($existingUser.AssignedLicenses.Count) licenses assigned"
        }

        if ($resultCode -eq 200) {
            # Create the new user with a random password
            $PasswordProfile = @{
                Password                      = -join ((65..90) + (97..122) + (48..57) + (33..47) | Get-Random -Count 16 | ForEach-Object { [char]$_ })
                ForceChangePasswordNextSignIn = $true
            }

            # Create the new user
            try {
                Write-Host "Attempting to create user: $newUserUpn"

                $userParams = @{
                    UserPrincipalName = $newUserUpn
                    DisplayName       = $NewUserDisplayName
                    GivenName         = $NewUserFirstName
                    Surname           = $NewUserLastName
                    MailNickname      = ($NewUserEmail.Split('@')[0])
                    AccountEnabled    = $true
                    PasswordProfile   = $PasswordProfile
                    UsageLocation     = "AU"
                }

                $newUser = New-MgUser -BodyParameter $userParams -ErrorAction Stop
                Write-Host "User created with ID: $($newUser.Id)"
            }
            catch {
                Write-Host "ERROR creating user: $_"
                $message = "Failed to create user: $_"
                $resultCode = 500
            }

            # Only proceed with licenses and groups if user was created successfully
            if ($newUser -and $newUser.Id -and $resultCode -eq 200) {
                # Assign the same licenses as the existing user
                if ($existingUser.AssignedLicenses.Count -gt 0) {
                    Write-Host "Assigning licenses..."
                    $licensesAssigned = 0
                    $licensesSkipped = 0

                    $existingUser.AssignedLicenses | ForEach-Object {
                        try {
                            Set-MgUserLicense -UserId $newUser.Id -AddLicenses @{ SkuId = $_.SkuId } -RemoveLicenses @() -ErrorAction Stop
                            $licensesAssigned++
                            Write-Host "Assigned license: $($_.SkuId)"
                        }
                        catch {
                            $licensesSkipped++
                            Write-Host "WARNING: Could not assign license $($_.SkuId) (may not be available): $($_.Exception.Message)"
                        }
                    }
                    Write-Host "Licenses summary: $licensesAssigned assigned successfully, $licensesSkipped skipped"
                }

                # Get the groups the existing user is a member of
                $existingUserGroups = Get-MgUserMemberOf -UserId $existingUser.Id

                # Add the new user to the same groups
                if ($existingUserGroups.Count -gt 0) {
                    Write-Host "Adding to groups..."
                    $securityGroupsAdded = 0
                    $mailEnabledGroupsAdded = 0
                    $groupsFailed = 0

                    # Connect to Exchange Online for mail-enabled groups using certificate authentication
                    try {
                        Write-Host "Connecting to Exchange Online with certificate..."
                        Write-Host "Tenant ID: $TenantId"
                        Write-Host "App ID: $($env:Ms365_AuthAppId)"
                        Write-Host "Cert Thumbprint: $($env:Ms365_CertThumbprint)"

                        # Use Tenant ID directly (Exchange accepts GUID format)
                        Connect-ExchangeOnline -AppId $env:Ms365_AuthAppId -CertificateThumbprint $env:Ms365_CertThumbprint -Organization $TenantId -ShowBanner:$false -ErrorAction Stop
                        $exchangeConnected = $true
                        Write-Host "Exchange Online connected successfully"
                    }
                    catch {
                        Write-Host "WARNING: Could not connect to Exchange Online: $_"
                        Write-Host "Mail-enabled groups and distribution lists will be skipped."
                        $exchangeConnected = $false
                    }

                    $existingUserGroups | ForEach-Object {
                        if ($_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
                            $groupId = $_.Id

                            # Try Graph API first (works for security groups)
                            try {
                                $params = @{
                                    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($newUser.Id)"
                                }
                                New-MgGroupMember -GroupId $groupId -BodyParameter $params -ErrorAction Stop
                                $securityGroupsAdded++
                                Write-Host "Added to security group: $groupId"
                            }
                            catch {
                                $errorMessage = $_.Exception.Message

                                # If it's a mail-enabled group error, try Exchange Online
                                if (($errorMessage -like "*mail-enabled*" -or $errorMessage -like "*distribution list*") -and $exchangeConnected) {
                                    try {
                                        # Get group details to find the name/identity
                                        $group = Get-MgGroup -GroupId $groupId -Property Id,DisplayName,Mail,MailEnabled
                                        $groupIdentity = if ($group.Mail) { $group.Mail } else { $group.DisplayName }

                                        Add-DistributionGroupMember -Identity $groupIdentity -Member $newUserUpn -ErrorAction Stop
                                        $mailEnabledGroupsAdded++
                                        Write-Host "Added to mail-enabled group: $groupIdentity"
                                    }
                                    catch {
                                        Write-Host "ERROR: Could not add to mail-enabled group $groupIdentity : $_"
                                        $groupsFailed++
                                    }
                                }
                                else {
                                    Write-Host "ERROR: Could not add to group $groupId : $errorMessage"
                                    $groupsFailed++
                                }
                            }
                        }
                    }

                    # Disconnect Exchange Online
                    if ($exchangeConnected) {
                        try {
                            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                        }
                        catch {
                            # Ignore disconnect errors
                        }
                    }

                    Write-Host "Groups summary: $securityGroupsAdded security groups added, $mailEnabledGroupsAdded mail-enabled groups added, $groupsFailed failed"
                }

                $message = "New user `"$NewUserEmail`" created successfully and assigned licenses and added to groups like user `"$ExistingUserEmail`"."
            }
        }
    }
}

$body = @{
    Message      = $message
    TicketId     = $TicketId
    ResultCode   = $resultCode
    ResultStatus = if ($resultCode -eq 200) { "Success" } else { "Failure" }
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode  = [HttpStatusCode]::OK
        Body        = $body
        ContentType = "application/json"
    })
