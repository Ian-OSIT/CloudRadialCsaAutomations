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

    # Retrieve the existing user's details
    $existingUser = Get-MgUser -UserPrincipalName $existingUserUpn

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
            # Retrieve all available licenses
            $availableLicenses = Get-MgSubscribedSku

            # Check if the available licenses match the existing user's licenses
            $existingUserLicenseIds = $existingUser.AssignedLicenses.SkuId
            $missingLicenses = $availableLicenses | Where-Object { $existingUserLicenseIds -notcontains $_.SkuId }

            if ($missingLicenses.Count -gt 0) {
                Write-Host "The following licenses are missing for the new user:"
                $missingLicenses | ForEach-Object {
                    Write-Host "- $($_.SkuPartNumber)"
                }
                $resultCode = 500
                $message = "Request failed. The existing user `"$ExistingUserEmail`" has licenses that are not available for the new user."
            }
        }

        if ($resultCode -eq 200) {
            # Create the new user with a random password
            $PasswordProfile = @{
                Password                      = -join ((65..90) + (97..122) + (48..57) + (33..47) | Get-Random -Count 16 | ForEach-Object { [char]$_ })
                ForceChangePasswordNextSignIn = $true
            }

            # Create the new user
            $newUser = New-MgUser -UserPrincipalName $newUserUpn -DisplayName $NewUserDisplayName -GivenName $NewUserFirstName -Surname $NewUserLastName -MailNickname ($NewUserEmail.Split('@')[0]) -AccountEnabled $true -PasswordProfile $PasswordProfile -UsageLocation "AU"

            # Assign the same licenses as the existing user
            if ($existingUser.AssignedLicenses.Count -gt 0) {
                $existingUser.AssignedLicenses | ForEach-Object {
                    Set-MgUserLicense -UserId $newUser.Id -AddLicenses @{ SkuId = $_.SkuId } -RemoveLicenses @()
                }
            }

            # Get the groups the existing user is a member of
            $existingUserGroups = Get-MgUserMemberOf -UserId $existingUser.Id

            # Add the new user to the same groups
            if ($existingUserGroups.Count -gt 0) {
                $existingUserGroups | ForEach-Object {
                    if ($_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
                        New-MgGroupMember -GroupId $_.Id -DirectoryObjectId $newUser.Id
                    }
                }
            }

            $message = "New user `"$NewUserEmail`" created successfully and assigned licenses and added to groups like user `"$ExistingUserEmail`"."
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
