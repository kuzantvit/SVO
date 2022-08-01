Get-Content -Path C:\PS\GetFlash.txt | foreach {
    if (Test-Connection -ComputerName $_ -Count 1 -Quiet)
    {
        $flashFile  = (Get-ChildItem "\\$_\C`$\Windows\System32\Macromed\Flash\Flash*.ocx" -ErrorAction SilentlyContinue).FullName
        if ($flashFile -ne $null) 
        {
            Write-Verbose "Getting $flashFile from $_"
            [system.diagnostics.fileversioninfo]::GetVersionInfo($flashFile)
        }
    } 
} | Export-CSV -Path C:\PS\GetFlash.csv -NoTypeInformation