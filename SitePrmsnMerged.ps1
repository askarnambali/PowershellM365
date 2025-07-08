# --- CONFIGURATION ---
$tenantId           = "d4b07abe-b466-4558-a63f-dc15f05f3693"                     # e.g., 8f123456-xxxx-4f5c-xxxx-xxxxxxxxxxxx
$clientId           = "099a7e49-1e07-4b26-b592-daa79079711d"                     # App Registration's Application (client) ID
$certificatePath    = "C:\Users\askar.mohamed\OneDrive - AL BAYARI\Documents\PnP Certificate\BayariICT1.pfx"
$certificatePassword = "20Marza25$"
$adminSiteUrl       = "https://dandbdubai-admin.sharepoint.com"
$exportPath         = "C:\Users\askar.mohamed\OneDrive - AL BAYARI\Documents\VSCode Powershell\Sitespermission.csv"

# --- Connect using App-Only Certificate Authentication ---
Connect-PnPOnline -Url $adminSiteUrl `
                  -ClientId $clientId `
                  -Tenant $tenantId `
                  -CertificatePath $certificatePath `
                  -CertificatePassword (ConvertTo-SecureString $certificatePassword -AsPlainText -Force)

# --- Get All SharePoint Sites (excluding OneDrive) ---
$sites = Get-PnPTenantSite -IncludeOneDriveSites:$false | Where-Object { $_.Url -like "https://dandbdubai.sharepoint.com/sites/*" }

# --- Prepare Output ---
$report = @()

foreach ($site in $sites) {
    try {
        Write-Host "🔍 Processing: $($site.Url)" -ForegroundColor Cyan

        $ownersList   = "N/A"
        $membersList  = "N/A"
        $visitorsList = "N/A"

        # If site is group-connected (has a valid GroupId)
        if ($site.GroupId -and $site.GroupId -ne [Guid]::Empty) {
            $groupId = $site.GroupId.ToString()
            try {
                # --- Get Owners via Graph ---
                $ownersResponse = Invoke-PnPGraphMethod -Url "groups/$groupId/owners" -Method GET
                $owners = $ownersResponse.value | Select-Object -ExpandProperty displayName
                if ($owners) { $ownersList = $owners -join ", " }

                # --- Get Members via Graph ---
                $membersResponse = Invoke-PnPGraphMethod -Url "groups/$groupId/members" -Method GET
                $members = $membersResponse.value | Select-Object -ExpandProperty displayName
                if ($members) { $membersList = $members -join ", " }

                $visitorsList = "N/A"  # No visitors group for group-connected sites
            }
            catch {
                Write-Warning "⚠️ Failed to fetch group owners/members for: $($site.Url)"
            }
        }
        else {
            # --- Classic site: Use SharePoint groups ---
            Connect-PnPOnline -Url $site.Url `
                              -ClientId $clientId `
                              -Tenant $tenantId `
                              -CertificatePath $certificatePath `
                              -CertificatePassword (ConvertTo-SecureString $certificatePassword -AsPlainText -Force)

            $groups = Get-PnPGroup
            foreach ($group in $groups) {
                $members = (Get-PnPGroupMember -Identity $group | Select-Object -ExpandProperty UserPrincipalName) -join ", "
                switch -Wildcard ($group.Title) {
                    "*Owners*"   { $ownersList   = $members }
                    "*Members*"  { $membersList  = $members }
                    "*Visitors*" { $visitorsList = $members }
                }
            }
        }

           # === Get file count from "Documents" library ===
        try {
            $fileCount = (Get-PnPListItem -List "Documents" -PageSize 1000 -Fields "FileLeafRef").Count
        } catch {
            $fileCount = "Error"
            Write-Warning "⚠️ Could not get file count for $siteUrl"
        }

        $report += [PSCustomObject]@{
            SiteTitle = $site.Title
            SiteUrl   = $site.Url
            Owners    = $ownersList
            Members   = $membersList
            Visitors  = $visitorsList
            FileCount = $fileCount
        }
    }
    catch {
        Write-Warning "⚠️ Failed to process $($site.Url): $_"
    }
}

# --- Export to CSV ---
$report | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
Write-Host "`n✅ Report exported to: $exportPath" -ForegroundColor Green
