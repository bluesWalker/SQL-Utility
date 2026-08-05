$script:Failures = 0

function Assert-True([bool] $Condition, [string] $Message) {
    if (-not $Condition) {
        $script:Failures++
        Write-Error -ErrorAction Continue "FAIL: $Message"
    }
}

function Assert-Equal($Expected, $Actual, [string] $Message) {
    Assert-True ($Expected -eq $Actual) "$Message (expected '$Expected', got '$Actual')"
}

function Assert-Throws([scriptblock] $Action, [string] $ExceptionType, [string] $Message) {
    $caught = $null
    try { & $Action } catch { $caught = $_.Exception }
    Assert-True ($null -ne $caught) $Message
    if ($null -ne $caught -and $ExceptionType) {
        Assert-Equal $ExceptionType $caught.GetType().FullName "$Message exception type"
    }
}

function Complete-TestFile([string] $SuccessMessage) {
    if ($script:Failures -gt 0) { exit 1 }
    Write-Host $SuccessMessage
}
