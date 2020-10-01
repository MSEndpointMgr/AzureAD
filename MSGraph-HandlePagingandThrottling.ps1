# Scriptmethod created by: Jan Ketil Skanke 
# Twitter @JankeSkanke
# https://msendpointmgr.com 
# Create result object before DO. 
$QueryResults = @()
$Uri = 'https://graph.microsoft.com/beta/users'
# Invoke REST method and fetch data until there are no pages left.
do {
    $RetryIn = "0"
    $ThrottledRun = $false  
    Write-Output "Querying $Uri..." 
    try{
        $Results = Invoke-RestMethod -Method Get -Uri $Uri -ContentType "application/json" -Headers $Header -ErrorAction Continue
    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $Myerror = $_.Exception
        if (($Myerror.Response.StatusCode) -eq "429"){
            $ThrottledRun = $true
            $RetryIn = $Myerror.Response.Headers["Retry-After"] 
            Write-Warning -Message "Graph queries is being throttled"
            Write-Output "Setting throttle retry to $($RetryIn) seconds"
        }else
        {
            Write-Error -Message "Inital graph query failed with message: $ErrorMessage"
            Exit 1
        }
    } 
    # Check if run is throttled, and skip adding data and keep current URI for retry
    if ($ThrottledRun -eq $false){
        #If request is not throttled put data into result object
        $QueryResults += $Results.value
        #If request is not trottled, go to nextlink if available to fetch more data
        $uri = $Results.'@odata.nextlink'
    }
    Start-Sleep -Seconds $RetryIn
} until (!($uri))
