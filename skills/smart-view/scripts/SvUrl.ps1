# Smart View URL of an EPM Cloud (Planning-based) instance: https://<host>/HyperionPlanning/SmartView.
# The base URL of the instance, /epmcloud or a browser URL answer "Unable to connect to the
# provider" or "Empty response from provider". Other hosts (Essbase, on-premises) are left as given.
function ConvertTo-SvUrl([string] $Url) {
    $u = $Url.Trim()
    if ($u -notmatch '^[a-z]+://') { $u = "https://$u" }
    $uri = [uri]$u
    if ($uri.Host -match '\.oraclecloud\.com$' -and $uri.AbsolutePath -notmatch '^/HyperionPlanning/SmartView/?$') {
        return "https://$($uri.Host)/HyperionPlanning/SmartView"
    }
    return $u.TrimEnd('/')
}
