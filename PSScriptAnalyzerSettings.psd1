@{
    Severity = @('Error', 'Warning', 'Information')

    Rules    = @{
        # Keep the scripts runnable on Windows PowerShell 5.1 as well as PowerShell 7.
        PSUseCompatibleSyntax = @{
            Enable         = $true
            TargetVersions = @('5.1', '7.0')
        }
    }
}
