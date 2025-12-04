<#
.SYNOPSIS
Gets the permanent, location-independent URL for a SharePoint Online document using its Document ID.

.DESCRIPTION
This script connects to SharePoint Online, extracts the list item details for a given document URL 
(handling modern sharing links heuristically), and retrieves the Document ID (DlcDocId) and its 
corresponding permanent redirect URL (DlcDocIdUrl) via the SharePoint REST API.

.NOTES
PnP.PowerShell Module: This script requires the PnP.PowerShell module to be installed.
Authentication: It uses Interactive login, which typically defaults to the Device Code flow on Mac/Linux 
or opens a browser window on Windows.

.ENTRA_ID_SETUP (CRITICAL for 'App is not allowed' errors)
If you encounter an "App is not allowed to call SPO" error (meaning the public PnP Client ID is blocked), 
you MUST use a custom Entra ID App Registration created within your tenant. Follow these steps:

1. Register App: Go to the Entra ID Portal -> App registrations -> New registration.
   - Name: PnP Document ID Lookup (or similar).
   - Supported account types: Accounts in this organizational directory only.
   - Redirect URI: Select 'Public client/native (mobile & desktop)' and enter 'http://localhost'.

2. Configure Permissions: Go to API permissions -> Add a permission -> SharePoint -> Delegated permissions.
   - Select: Sites.FullControl.All (or Sites.Read.All for read-only access).
   - ACTION: An admin MUST click 'Grant admin consent for [your tenant name]'.

3. Update Script: Copy the Application (Client) ID from the app's Overview page and replace the value of 
   $CustomClientID below with your new ID.

Example Manual Connection Command:
Connect-PnPOnline -Url https://ribash.sharepoint.com/s/RIBASH-ICF-BEIS -Interactive -ClientID e8b2afe3-e580-4e46-8b24-ded9c8a99042
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$DocumentUrl,

    [string]$TenantUrl = "" # Optional: If not provided, it will be inferred from the document URL
)

function Get-SharePointPermanentUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Url,

        [string]$TenantBaseUrl
    )

    # 1. Determine the Site URL and Server-Relative URL
    Write-Host "Analyzing URL: $Url" -ForegroundColor Yellow

    # Extract the SharePoint Online domain/tenant base URL
    if (-not $TenantBaseUrl) {
        $Uri = New-Object System.Uri($Url)
        $TenantBaseUrl = "$($Uri.Scheme)://$($Uri.Host)"
    }

    # Extract the server-relative path (e.g., /sites/mysite/Shared%20Documents/doc.docx)
    # Note: For modern links, the path extracted here is often the encoded sharing path, 
    # but we rely on the PnP connection context to resolve the actual file when calling the REST API.
    $CleanedUrl = $Url.Split('?')[0]
    $ServerRelativeUrl = $CleanedUrl.Substring($TenantBaseUrl.Length)
    
    Write-Host "Tenant URL found: $TenantBaseUrl"
    Write-Host "Relative Path used for API context: $ServerRelativeUrl"
    
    # 2. Check for Active Connection
    # We rely on the connection established in the main execution block.
    try {
        $Connection = Get-PnPConnection -ErrorAction Stop
        if (-not $Connection) {
            Write-Error "Connection is missing. Ensure Connect-PnPOnline was executed successfully."
            return
        }
    }
    catch {
        Write-Error "Could not verify PnP connection. $($_.Exception.Message)"
        return
    }

    # 3. Retrieve the Document ID metadata
    try {
        # Use GetFileByServerRelativeUrl and select properties from the ListItemAllFields
        # This function requires the correct server-relative path of the file.
        # We must ensure $ServerRelativeUrl is URL-encoded for the REST API call, but PnP generally handles simple paths.
        
        # We use a literal path for the REST endpoint construction, including $select for efficiency.
        $FileApiEndpoint = "/_api/web/GetFileByServerRelativeUrl('${ServerRelativeUrl}')/ListItemAllFields?`$select=FileRef,DlcDocId,DlcDocIdUrl"
        
        $ItemData = Invoke-PnPSPRestMethod -Url $FileApiEndpoint -Method Get -ErrorAction Stop
        
        # Check if Document ID exists
        $DocumentID = $ItemData.DlcDocId
        $DocumentIDUrl = $ItemData.DlcDocIdUrl

        if (-not $DocumentID) {
            Write-Warning "Document ID (DlcDocId) not found for this document. Ensure the Document ID Service feature is active on the site collection, or ensure the provided URL is the *canonical* path (e.g., .../Library/Document.docx) and not an encoded sharing link, as PnP may fail to resolve the correct server path."
            return
            }

        # 4. Extract and Format the Permanent URL
        # The DlcDocIdUrl field often contains the permanent URL in the format:
        # "https://tenant.sharepoint.com/sites/mysite/_layouts/15/DocIdRedir.aspx?ID=DOC-1234567890-1, DOC-1234567890-1"
        # We clean this up by taking the first element of the split.

        $PermanentUrl = $DocumentIDUrl.Split(',')[0].Trim()
        
        Write-Host ""
        Write-Host "-------------------------------------------"
        Write-Host "Document ID Found: $($DocumentID)" -ForegroundColor Green
        Write-Host "Permanent Redirect URL:" -ForegroundColor Green
        Write-Host "$PermanentUrl" -ForegroundColor Cyan
        Write-Host "-------------------------------------------"
        return $PermanentUrl

    }
    catch {
        Write-Error "An error occurred during the API call or file lookup. This may happen if the URL provided is a highly encoded sharing link, or if the Document ID Service is not active."
        Write-Error "Original PowerShell Error: $($_.Exception.Message)"
    }
}

# --- Execution ---

# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# *** CRITICAL FIX FOR: invalid_request / App is not allowed to call SPO ***
# If you receive the "App is not allowed..." error, you MUST create your own Entra ID App Registration
# and update the Client ID below with the Application (Client) ID from your custom registration.
# The default ID below is the blocked PnP ID.
$CustomClientID = "e8b2afe3-e580-4e46-8b24-ded9c8a99042" 
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# 1. Check for PnP Module (optional but helpful)
if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
    Write-Warning "PnP.PowerShell module is not installed. Please install it using: Install-Module PnP.PowerShell -Scope CurrentUser"
    exit
}

# 2. Prompt for input if parameters are missing
if (-not $DocumentUrl) {
    $DocumentUrl = Read-Host "Enter the full SharePoint Document URL (e.g., https://tenant.sharepoint.com/sites/site/Library/doc.docx)"
    if (-not $DocumentUrl) {
        Write-Error "Document URL cannot be empty."
        exit
    }
}

# 3. Determine the Site URL for connection
try {
    # Extract the site URL for connection (this is a heuristic and might need adjustment for complex paths)
    $Uri = New-Object System.Uri($DocumentUrl)
    $TenantHost = "$($Uri.Scheme)://$($Uri.Host)"
    $SiteUrlGuess = $TenantHost

    # Check for Modern Sharing Link format (/:x:/s/SITE_NAME/...)
    if ($DocumentUrl -match "(^https?://[^/]+)/:x:/s/([^/]+)") {
        # Site URL is Host/s/SiteName. This is the fix for the user's provided URL format.
        $SiteUrlGuess = "$($Matches[1])/s/$($Matches[2])"
    }
    # Check for standard site collection path (e.g., /sites/SITE_NAME/...)
    elseif ($DocumentUrl -match "(^https?://[^/]+)/sites/([^/]+)") {
        # Site URL is Host/sites/SiteName
        $SiteUrlGuess = "$($Matches[1])/sites/$($Matches[2])"
    }
    # Fallback to tenant root if no specific site path is found
    # We rely on the site-name being correct for connection.

} catch {
    Write-Error "Could not parse the site URL from the provided document URL. Please check the format. $($_.Exception.Message)"
    exit
}

# 4. Connect to SharePoint (This is where the pre-authentication happens)
try {
    # Check if a connection already exists to avoid unnecessary prompts
    $ExistingConnection = Get-PnPConnection -ErrorAction SilentlyContinue
    
    if (-not $ExistingConnection) {
        
        $ClientIDToUse = $CustomClientID
        
        Write-Host "Attempting to connect to SharePoint Online site: $SiteUrlGuess using Interactive Authentication." -ForegroundColor Green
        Write-Host "NOTE: If this fails, you MUST replace '$CustomClientID' with your custom Entra ID App ID." -ForegroundColor Red
        Write-Host "Please follow the instructions below to sign in via your web browser (a code and URL should appear):" -ForegroundColor Green
        
        # Use the explicit Client ID for connection
        Connect-PnPOnline -Url $SiteUrlGuess -Interactive -ClientID $ClientIDToUse -ErrorAction Stop
        Write-Host "Successfully connected to $SiteUrlGuess using Interactive flow." -ForegroundColor Green
    } else {
        # If a connection exists, check if it's the right tenant.
        $CurrentConnectionUrl = $ExistingConnection.Url
        # Simple check: Does the host name match?
        if ($CurrentConnectionUrl -notmatch $SiteUrlGuess.Split('/')[2]) {
            Write-Host "Warning: An existing connection to '$CurrentConnectionUrl' was found, but the document belongs to a different host/site: '$SiteUrlGuess'." -ForegroundColor Red
            Write-Host "ACTION REQUIRED: Please run 'Disconnect-PnPOnline' and re-run this script to sign into the correct tenant/site." -ForegroundColor Red
            # Force an exit here to prevent the script from trying to use the wrong connection.
            exit
        } else {
            Write-Host "Reusing existing connection to $CurrentConnectionUrl." -ForegroundColor Green
        }
    }
    
    # 5. Call the function to get the permanent URL
    Get-SharePointPermanentUrl -Url $DocumentUrl -TenantBaseUrl $TenantUrl

} catch {
    # Improved error message for connectivity issues.
    Write-Error "CONNECTION FAILURE: Failed to establish a SharePoint session."
    Write-Error "The script attempted to connect to '$SiteUrlGuess'."
    Write-Error "Please verify the following actions to resolve this issue:"
    Write-Error "1. **Check Session State:** Run the command 'Disconnect-PnPOnline' and then re-run this script to ensure a fresh authentication attempt."
    Write-Error "2. **ENTRA ID BLOCK:** You MUST create and use a custom Entra ID Application ID to proceed."
    Write-Error "Original PowerShell Error: $($_.Exception.Message)"
}
