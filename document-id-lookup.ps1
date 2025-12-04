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
    $ServerRelativeUrl = $Url.Substring($TenantBaseUrl.Length).Split('?')[0]
    $ServerRelativeUrl = $ServerRelativeUrl -replace '%20', ' ' # Decode spaces for API call

    Write-Host "Tenant URL found: $TenantBaseUrl"
    Write-Host "Relative Path: $ServerRelativeUrl"

    # The Document ID is stored on the document's list item.
    # The SiteUrl extraction logic is handled in the main execution block for connection.
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
        if ($Connection.Url -ne $SiteUrl) {
             Write-Host "Note: Current PnP connection is to '$($Connection.Url)', but document is on '$SiteUrl'." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Could not verify PnP connection. $($_.Exception.Message)"
        return
    }

    # 3. Retrieve the Document ID metadata
    try {
        # Use GetFileByServerRelativeUrl and select properties from the ListItemAllFields
        # Note: The ServerRelativeUrl already contains the site path if it's not the root site.
        $FileApiEndpoint = "/_api/web/GetFileByServerRelativeUrl('${ServerRelativeUrl}')/ListItemAllFields"
        
        $ItemData = Invoke-PnPSPRestMethod -Url $FileApiEndpoint -Method Get -Select "FileRef, DlcDocId, DlcDocIdUrl" -ErrorAction Stop
        
        # Check if Document ID exists
        $DocumentID = $ItemData.DlcDocId
        $DocumentIDUrl = $ItemData.DlcDocIdUrl

        if (-not $DocumentID) {
            Write-Warning "Document ID (DlcDocId) not found for this document. Ensure the Document ID Service feature is active on the site collection."
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
    $SiteUrlGuess = $DocumentUrl.Split("/Shared Documents")[0] 
    if ($SiteUrlGuess -notlike "*sites/*") {
        # Attempt to find the site base URL if it's not a /sites/ path
        $SiteUrlGuess = $DocumentUrl.Split("/Lists/")[0]
        if ($SiteUrlGuess -notlike "*sites/*") {
            # Last resort: use the tenant root
            $Uri = New-Object System.Uri($DocumentUrl)
            $SiteUrlGuess = "$($Uri.Scheme)://$($Uri.Host)"
        }
    }
} catch {
    Write-Error "Could not parse the site URL from the provided document URL. Please check the format."
    exit
}

# 4. Connect to SharePoint (This is where the pre-authentication happens)
try {
    # Check if a connection already exists to avoid unnecessary prompts
    if (-not (Get-PnPConnection -ErrorAction SilentlyContinue)) {
        
        Write-Host "Attempting to connect to SharePoint Online site: $SiteUrlGuess using Device Code Authentication." -ForegroundColor Green
        Write-Host "Please follow the instructions that appear shortly to sign in." -ForegroundColor Green
        
        # Use DeviceAuth to bypass Client ID issues. This will print a code to the console.
        Connect-PnPOnline -Url $SiteUrlGuess -DeviceAuth -ErrorAction Stop
        Write-Host "Successfully connected to $SiteUrlGuess using Device Code flow." -ForegroundColor Green
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
