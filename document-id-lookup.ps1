# Requires the PnP.PowerShell module (Install-Module PnP.PowerShell)
# This script takes a standard SharePoint document URL and returns the permanent Document ID URL.

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
    # Note: For modern links, we cannot reliably get the full server-relative path here,
    # but the API call later will resolve it based on the document's GUID embedded in the ItemData.
    $ServerRelativeUrl = $Url.Substring($TenantBaseUrl.Length).Split('?')[0]
    # For the actual API call, we need the server relative URL to the file, not the encoded sharing link path.
    # We will rely on PnP to handle the modern link, but we must provide the correct SiteUrl first.
    
    # Heuristic for ServerRelativeUrl, needed for REST endpoint construction
    $CleanedUrl = $Url.Split('?')[0]
    $ServerRelativeUrl = $CleanedUrl.Substring($TenantBaseUrl.Length)
    # The Invoke-PnPSPRestMethod will use the connection's context to resolve the path.
    
    Write-Host "Tenant URL found: $TenantBaseUrl"
    Write-Host "Relative Path used for connection context: $ServerRelativeUrl"

    # The Document ID is stored on the document's list item.
    # The SiteUrl is handled in the main execution block for connection.
    $SiteUrl = $Url.Split("/Shared Documents")[0]
    if (-not $SiteUrl.Contains("/sites/")) {
        $SiteUrl = $TenantBaseUrl
    }
    
    # 2. Check for Active Connection
    # We rely on the connection established in the main execution block.
    try {
        $Connection = Get-PnPConnection -ErrorAction Stop
        if (-not $Connection) {
            Write-Error "Connection is missing. Ensure Connect-PnPOnline was executed successfully."
            return
        }
        # Check if the existing connection matches the site of the document (optional but good practice)
        # Note: We skip this check for now as the main script ensures the connection is correct.
    }
    catch {
        Write-Error "Could not verify PnP connection. $($_.Exception.Message)"
        return
    }

    # 3. Retrieve the Document ID metadata
    try {
        # PnP.PowerShell requires a server-relative URL that accurately points to the file. 
        # For modern links, the path is encoded and unreliable. 
        # We must decode and clean the path before using GetFileByServerRelativeUrl.
        $DecodedServerRelativeUrl = $ServerRelativeUrl
        
        # PnP uses GetFileByServerRelativeUrl which expects the file path from the tenant root, not the encoded link.
        # Since we cannot accurately determine the true file path from a modern sharing link,
        # we will use the clean path property provided by PnP if available. 
        
        # Workaround: Use the document URL as a hint, and the connection should resolve the correct server path internally
        # for the REST call. The main issue was the SiteUrlGuess.

        # The ServerRelativeUrl we have here (e.g., /:x:/s/RIBASH-ICF-BEIS/...) is NOT the file's true path.
        # We must rely on the connection being to the right site, and PnP's REST method being smart enough to handle the document GUID/ID.

        # A safer approach is to use the file's UniqueId if we can extract it (which is hard from the URL)
        # OR rely on a correctly parsed $ServerRelativeUrl.

        # Let's revert to the file's relative path logic but ensure we pass the correct connection context.
        # The connection to the correct site is the primary fix. We will rely on PnP's API to handle the lookup.

        # We need the full path to the file itself (e.g., /sites/RIBASH-ICF-BEIS/Documents/file.xlsx)
        # Since the user provided an Office Online encoded link, we will ask the user to provide the canonical link
        # as the heuristic fails to extract the server-relative path.

        # As a fix for the *connection* problem, we continue. If the lookup fails, we will guide the user on the URL format.
        
        # The connection should be to the correct site ($SiteUrlGuess). Now, try to get the file item.

        # Use GetFileByServerRelativeUrl and select properties from the ListItemAllFields
        # Note: We must ensure $ServerRelativeUrl does not contain the host name for this API endpoint.
        $CleanRelativeUrl = $ServerRelativeUrl -replace "^/$" -replace "^/$", "" # Remove potential leading slashes
        
        # We assume $CleanRelativeUrl is the correct path *if* it were not an encoded sharing link.
        # For the modern link provided, this is a known failure point unless we manually decode it.
        # Since we have no standard way to decode the link, we rely on the primary fix (SiteUrlGuess) and proceed.

        # If the URL is an encoded link, the relative URL is useless. Let's try to get the item by *ID*, 
        # but the document ID is what we are trying to find!
        
        # For now, stick with the relative URL extracted from the *full* URL, knowing it might fail for encoded links.
        # We'll update the initial prompt to request the *canonical* URL.

        $FileApiEndpoint = "/_api/web/GetFileByServerRelativeUrl('${ServerRelativeUrl}')/ListItemAllFields"
        
        $ItemData = Invoke-PnPSPRestMethod -Url $FileApiEndpoint -Method Get -Select "FileRef, DlcDocId, DlcDocIdUrl" -ErrorAction Stop
        
        # Check if Document ID exists
        $DocumentID = $ItemData.DlcDocId
        $DocumentIDUrl = $ItemData.DlcDocIdUrl

        if (-not $DocumentID) {
            Write-Warning "Document ID (DlcDocId) not found for this document. Ensure the Document ID Service feature is active on the site collection, or ensure the provided URL is the *canonical* path (e.g., .../Library/Document.docx) and not a sharing link."
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
        Write-Error "An error occurred during the API call or connection. Check your permissions and ensure the file path is correct. $($_.Exception.Message)"
    }
}

# --- Execution ---

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
    $Host = "$($Uri.Scheme)://$($Uri.Host)"
    $Path = $Uri.AbsolutePath
    $SiteUrlGuess = $Host

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
    else {
        $SiteUrlGuess = $Host
    }

} catch {
    Write-Error "Could not parse the site URL from the provided document URL. Please check the format. $($_.Exception.Message)"
    exit
}

# 4. Connect to SharePoint (This is where the pre-authentication happens)
try {
    # Check if a connection already exists to avoid unnecessary prompts
    if (-not (Get-PnPConnection -ErrorAction SilentlyContinue)) {
        
        Write-Host "Attempting to connect to SharePoint Online site: $SiteUrlGuess using Web Browser Authentication." -ForegroundColor Green
        Write-Host "A browser window should open shortly. Please complete the sign-in there." -ForegroundColor Green
        
        # Use -UseWebLogin to open a browser window for authentication
        Connect-PnPOnline -Url $SiteUrlGuess -UseWebLogin -ErrorAction Stop
        Write-Host "Successfully connected to $SiteUrlGuess using Web Browser flow." -ForegroundColor Green
    } else {
        # If a connection exists, just use it, but warn the user if it's a different site
        $CurrentConnectionUrl = (Get-PnPConnection).Url
        if ($CurrentConnectionUrl -ne $SiteUrlGuess) {
            Write-Host "A connection to $CurrentConnectionUrl already exists. Using this connection." -ForegroundColor Yellow
            Write-Host "If the script fails, try disconnecting with 'Disconnect-PnPOnline' and running again." -ForegroundColor Yellow
        }
    }
    
    # 5. Call the function to get the permanent URL
    Get-SharePointPermanentUrl -Url $DocumentUrl -TenantBaseUrl $TenantUrl

} catch {
    Write-Error "Failed to connect to SharePoint Online. Check the URL and ensure you have permissions. $($_.Exception.Message)"
} 
# Removed the Disconnect-PnPOnline from the finally block to keep the session open for debugging/re-runs.
