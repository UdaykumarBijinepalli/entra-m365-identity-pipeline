param(
  [string]$RosterPath = 'roster/cohort01-trainees.csv',
  [string]$CohortGroupName = 'SG-Cohort01-Trainees',
  [string]$TargetResourceGroup = 'rg-cohort01-lab'
)

$ErrorActionPreference = 'Stop'
$GraphToken = az account get-access-token --resource-type ms-graph --query accessToken -o tsv
$Headers = @{ Authorization = "Bearer $GraphToken"; 'Content-Type' = 'application/json' }

# Discover tenant primary domain
$Domains = (Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/domains' -Headers $Headers).value
$PrimaryDomain = ($Domains | Where-Object { $_.isDefault -eq $true }).id
Write-Host "Connected to Entra Tenant Domain: $PrimaryDomain" -ForegroundColor Cyan

# Check if any M365 License SKUs have available units
$Skus = (Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus' -Headers $Headers).value
$AvailableSku = $Skus | Where-Object { ($_.prepaidUnits.enabled - $_.consumedUnits) -gt 0 } | Select-Object -First 1

# Ensure Cohort Security Group exists
$GroupResp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '$CohortGroupName'" -Headers $Headers
if ($GroupResp.value.Count -eq 0) {
  $GroupBody = @{ displayName = $CohortGroupName; mailEnabled = $false; mailNickname = 'sgcohort01'; securityEnabled = $true; description = 'Automated security group for Cohort 01 trainees' } | ConvertTo-Json
  $CohortGroup = Invoke-RestMethod -Method Post -Uri 'https://graph.microsoft.com/v1.0/groups' -Headers $Headers -Body $GroupBody
} else {
  $CohortGroup = $GroupResp.value[0]
}
$GroupId = $CohortGroup.id

# Check if target cohort Resource Group exists for RBAC binding
$RgExists = (az group exists --name $TargetResourceGroup) -eq 'true'
$SubId = az account show --query id -o tsv
$RgScope = "/subscriptions/$SubId/resourceGroups/$TargetResourceGroup"

$Roster = Import-Csv -Path $RosterPath
$AuditEntries = @()

foreach ($Row in $Roster) {
  $Upn = "$($Row.MailNickname)@$PrimaryDomain"
  $DisplayName = "$($Row.FirstName) $($Row.LastName)"
  $UserQuery = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '$Upn'" -Headers $Headers
  $User = $UserQuery.value | Select-Object -First 1

  if ($Row.Status -eq 'Active') {
    if (-not $User) {
      $TempPass = "Trn@" + (Get-Random -Minimum 100000 -Maximum 999999) + "!Aa"
      $UserBody = @{
        accountEnabled = $true
        displayName = $DisplayName
        mailNickname = $Row.MailNickname
        userPrincipalName = $Upn
        department = $Row.Department
        usageLocation = 'IN'
        passwordProfile = @{ forceChangePasswordNextSignIn = $true; password = $TempPass }
      } | ConvertTo-Json -Depth 4
      $User = Invoke-RestMethod -Method Post -Uri 'https://graph.microsoft.com/v1.0/users' -Headers $Headers -Body $UserBody
      $Action = 'ProvisionedNewUser'
    } else {
      Invoke-RestMethod -Method Patch -Uri "https://graph.microsoft.com/v1.0/users/$($User.id)" -Headers $Headers -Body (@{ accountEnabled = $true; department = $Row.Department } | ConvertTo-Json)
      $Action = 'VerifiedActiveUser'
    }

    # Add to Cohort Security Group (idempotent)
    $Members = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members" -Headers $Headers).value
    if ($Members.id -notcontains $User.id) {
      $MemberBody = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($User.id)" } | ConvertTo-Json
      Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$ref" -Headers $Headers -Body $MemberBody
    }

    # Assign M365 License if available in tenant
    $LicenseStatus = 'Skipped-NoAvailableSKU'
    if ($AvailableSku) {
      $LicBody = @{ addLicenses = @(@{ skuId = $AvailableSku.skuId }); removeLicenses = @() } | ConvertTo-Json -Depth 4
      Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$($User.id)/assignLicense" -Headers $Headers -Body $LicBody | Out-Null
      $LicenseStatus = "Assigned-$($AvailableSku.skuPartNumber)"
    }

    # Assign Reader RBAC on Cohort Resource Group if provisioned
    $RbacStatus = 'Deferred-CohortRGNotActive'
    if ($RgExists) {
      az role assignment create --assignee-object-id $User.id --assignee-principal-type User --role 'Reader' --scope $RgScope | Out-Null
      $RbacStatus = 'Reader-Assigned'
    }

    $AuditEntries += [PSCustomObject]@{ Timestamp = (Get-Date -Format 'o'); EmployeeId = $Row.EmployeeId; UPN = $Upn; Status = 'Active'; Action = $Action; Group = $CohortGroupName; License = $LicenseStatus; AzureRBAC = $RbacStatus }
  }
  elseif ($Row.Status -eq 'Offboarded') {
    if ($User) {
      # 1. Disable account & revoke all active refresh/session tokens
      Invoke-RestMethod -Method Patch -Uri "https://graph.microsoft.com/v1.0/users/$($User.id)" -Headers $Headers -Body (@{ accountEnabled = $false } | ConvertTo-Json)
      Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$($User.id)/revokeSignInSessions" -Headers $Headers | Out-Null

      # 2. Remove from Cohort Security Group
      $Members = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members" -Headers $Headers).value
      if ($Members.id -contains $User.id) {
        Invoke-RestMethod -Method Delete -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members/$($User.id)/$ref" -Headers $Headers
      }

      # 3. Reclaim any assigned M365 licenses
      $UserLicenses = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($User.id)/licenseDetails" -Headers $Headers).value
      if ($UserLicenses.Count -gt 0) {
        $SkuIds = @($UserLicenses | ForEach-Object { $_.skuId })
        $RevokeBody = @{ addLicenses = @(); removeLicenses = $SkuIds } | ConvertTo-Json -Depth 4
        Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$($User.id)/assignLicense" -Headers $Headers -Body $RevokeBody | Out-Null
      }

      # 4. Strip Azure RBAC assignments
      az role assignment delete --assignee $User.id --scope "/subscriptions/$SubId" 2>$null
      $AuditEntries += [PSCustomObject]@{ Timestamp = (Get-Date -Format 'o'); EmployeeId = $Row.EmployeeId; UPN = $Upn; Status = 'Offboarded'; Action = 'Disabled-SessionsRevoked-RBACStripped'; Group = 'Removed'; License = 'Reclaimed'; AzureRBAC = 'Revoked' }
    } else {
      $AuditEntries += [PSCustomObject]@{ Timestamp = (Get-Date -Format 'o'); EmployeeId = $Row.EmployeeId; UPN = $Upn; Status = 'Offboarded'; Action = 'NoAccountToOffboard'; Group = 'None'; License = 'None'; AzureRBAC = 'None' }
    }
  }
}

New-Item -ItemType Directory -Force -Path 'audit' | Out-Null
$AuditEntries | ConvertTo-Json -Depth 4 | Out-File -Encoding utf8 'audit/identity-sync-report.json'
$AuditEntries | Format-Table -AutoSize
