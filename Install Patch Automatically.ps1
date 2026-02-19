# CONFIGURATION
# Source server where patches exist
$SourceNode = "SourceServer"

# Full paths on source server
$Patch1SourcePath = "C:\Windows\ProPatches\Patches\windows10.0-kb5063877-x64-1809.msu"
$Patch2SourcePath = "C:\Windows\ProPatches\Patches\windows10.0-kb5068791-x64-1809.msu"

# Destination folder on target servers
$PatchFolder = "C:\Windows\ProPatches\Patches"
$DestinationShare = "C$\Windows\ProPatches\Patches"

# Server groups
$Patch1Servers = @("Server1","Server2","Server3")
$Patch2Servers = @("Server4","Server5","Server6")


# FUNCTION: COPY PATCH

function Copy-Patch {

    param (
        [string]$SourceFile,
        [array]$TargetServers
    )

    # Extract filename from full source path
    $FileName = Split-Path $SourceFile -Leaf

    foreach ($server in $TargetServers) {

        Write-Host "Copying $FileName to $server..."

        # Check if server is reachable
        if (Test-Path "\\$server\C$") {

            # Ensure patch folder exists on target
            Invoke-Command -ComputerName $server -ScriptBlock {
                param($PatchFolder)

                if (!(Test-Path $PatchFolder)) {
                    New-Item -Path $PatchFolder -ItemType Directory -Force | Out-Null
                }

            } -ArgumentList $PatchFolder

            # Copy directly from actual source file path
            Copy-Item -Path $SourceFile `
                      -Destination "\\$server\$DestinationShare\$FileName" `
                      -Force -Verbose
        }
        else {
            Write-Host "$server not reachable. Skipping." -ForegroundColor Red
        }
    }
}

# FUNCTION: INSTALL PATCH


function Install-Patch {
    param (
        [string]$KB,
        [array]$Servers
    )

    Invoke-Command -ComputerName $Servers -ThrottleLimit 32 -ScriptBlock {

        param($KB, $PatchFolder)

        $server = $env:COMPUTERNAME
        Write-Host "==== [$server] Processing $KB ===="

        # Skip if already installed
        if (Get-HotFix -Id $KB -ErrorAction SilentlyContinue) {
            Write-Host "[$server] $KB already installed. Skipping." -ForegroundColor Green
            return
        }

        # Find MSU
        $msu = Get-ChildItem -Path $PatchFolder -Filter "*$KB*.msu" -ErrorAction SilentlyContinue | Select-Object -First 1

        if (-not $msu) {
            Write-Host "[$server] Patch file not found." -ForegroundColor Red
            return
        }

        $taskName = "Install_$KB"

        $action = New-ScheduledTaskAction -Execute "wusa.exe" `
                  -Argument "`"$($msu.FullName)`" /quiet /norestart"

        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" `
                     -RunLevel Highest -LogonType ServiceAccount

        $task = New-ScheduledTask -Action $action -Principal $principal

        Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null

        Start-ScheduledTask -TaskName $taskName

        # Wait loop
        do {
            Start-Sleep -Seconds 10
            $state = (Get-ScheduledTask -TaskName $taskName).State
            Write-Host "[$server] Status: $state"
        } while ($state -eq "Running")

        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false

        # Final verification
        if (Get-HotFix -Id $KB -ErrorAction SilentlyContinue) {
            Write-Host "[$server] Install completed (no reboot)." -ForegroundColor Green
        }
        else {
            Write-Host "[$server] Install finished but not detected." -ForegroundColor Red
        }

    } -ArgumentList $KB, $PatchFolder
}

# EXECUTION FLOW


# Extract KB numbers from filenames
$KB1 = "KB5063877"
$KB2 = "KB5068791"

Write-Host "=== PATCH 1 PROCESS STARTED ===" -ForegroundColor Cyan
Copy-Patch -SourceFile $Patch1SourcePath -TargetServers $Patch1Servers
Install-Patch -KB $KB1 -Servers $Patch1Servers

Write-Host "=== PATCH 2 PROCESS STARTED ===" -ForegroundColor Cyan
Copy-Patch -SourceFile $Patch2SourcePath -TargetServers $Patch2Servers
Install-Patch -KB $KB2 -Servers $Patch2Servers

Write-Host "=== ALL OPERATIONS COMPLETED ===" -ForegroundColor Yellow
