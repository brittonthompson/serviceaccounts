<#
.NOTES
  The parameters are bound this way to allow parameter binding to the PS1 or use as a function. Additionally, PSBoundParamters allows the $MyInvocation options to work for dynamic path discoveries.
#>
[CmdletBinding()]
param
(
  [Parameter(ValueFromPipelineByPropertyName = $true, Position = 0)]
  [Alias("CN", "Computers", "Servers")]
  [string[]]$ComputerName = @($env:COMPUTERNAME),
  [string]$CSVPath,
  [switch]$Transcript
)


function Get-ServiceAccounts {
  <#
  .SYNOPSIS
      Collect logins for services and scheduled tasks that are likely to be service accounts
  .DESCRIPTION
      Remotely collect a CSV and log of all accounts associated with scheduled tasks and services on servers. Useful for when you plan to make changes to old admin passowrds.
  .PARAMETER ComputerName
      String array of computer names to check for service accounts
  .PARAMETER CSVPath
      String to change the default path of the CSV output
  .PARAMETER Transcript
      Switch to enable the detailed logging to a transcript file within the script directory
  .NOTES
      Britt Thompson
      bthompson@1path.com
  #>
  [CmdletBinding()]
  param
  (
    [Parameter(ValueFromPipelineByPropertyName = $true, Position = 0)]
    [Alias("CN", "Computers", "Servers")]
    [string[]]$ComputerName,
    [string]$CSVPath,
    [switch]$Transcript
  )
  Begin {
    # Establish script path for log output
    $OSVersion = [decimal]([environment]::OSVersion.Version.Major, [environment]::OSVersion.Version.Minor -join ".")
    #If Visual Stuido code, PowerShell ISE or PowerShell
    if ($psEditor) {
      $Context = Get-Item ($psEditor.GetEditorContext()).CurrentFile.Path -ErrorAction SilentlyContinue
    }
    elseif ($psISE) {
      $Context = Get-Item $psISE.CurrentFile.FullPath -ErrorAction SilentlyContinue
    }
    else {
      $Context = Get-Item $MyInvocation.ScriptName -ErrorAction SilentlyContinue
    }

    $ScriptPath = $Context.DirectoryName
    #$ScriptFullName = $Context.FullName
    $ScriptName = $Context.BaseName
    [string]$DateFormat = Get-Date -Format "yyyy-MM-dd_HHmm"
    $LogPath = "$ScriptPath\Logs"
    $LogFile = "$LogPath\$ScriptName`_$DateFormat.txt"
    if ($Transcript -and $LogFile) { Start-Transcript -Path $LogFile }
    if (-not (Test-Path $LogPath)) { New-Item -Path $LogPath -ItemType Directory -Force | Out-Null }
    if (-not $CSVPath) { $CSVPath = $LogPath }
    $CSVFile = "$CSVPath\$ScriptName`_$DateFormat.csv"
    $ExcludeTasks = @("Adobe Acrobat Update Task", "G2MUpdateTask", "G2MUploadTask", "Optimize Start Menu Cache", "User_Feed_Synchronization", "GoogleUpdate", "TaskName", "OneDrive")
    $Results = @()
  }
  Process {
    Write-Host @"

============================================================================
  Onepath Service Account Collection $(Get-Date)
============================================================================

"@
    # Loop through all computers
    foreach ($C in $ComputerName) {
      Write-Host "[$(Get-Date)] =============================== [ Server: $C ]" -ForegroundColor Green

      $Properties = @{ ClassName = "Win32_Service" }

      if ($C -ne $env:COMPUTERNAME) {
        # Test the availability of the server to be checked            
        $Avail = Test-Connection -ComputerName $C -Count 1 -ErrorAction 0

        $WSMan = Test-WSMan $C -ErrorAction SilentlyContinue
        if ($WSMan -and $WSMan.ProductVersion -match "3.0") {
          $Properties["CimSession"] = New-CimSession $C
        }
        else {
          $Properties["ComputerName"] = $C
        }
      }
      else { $Avail = $true }

      if ($Avail) {
        $Lang = "service accounts"
        Write-Host "[$(Get-Date)] Checking for $Lang"

        #If this machine is at least 2012 try using CIM first
        try {
          $Services = if ($OSVersion -ge 6.2) { 
            Get-CIMInstance @Properties
          }
          else {
            Get-WmiObject @Properties
          }
        }
        catch { 
          try { 
            #Just in case there was an issue with CIM, lets try with WMI
            if ($OSVersion -ge 6.2) { 
              $Services = Get-WmiObject @Properties
            }
            else {
              throw "Error using the Get-WMIObject command"
            }
          }
          catch {
            Write-Host "[$(Get-Date)] $($_.Exception.Message)" -ForegroundColor Red
          }
        }

        # If services are collected output to screen, log and results for CSV output
        if ($Services) {
          $Services = $Services | Where-Object { $_.StartName -ne "LocalSystem" -and $_.StartName -notlike "NT *" -and $_.StartName.Length -gt 1 } | 
          Select-Object @{Name = "ComputerName"; Expression = { $C } }, Name, StartName, StartMode, State, 
          #Create columns with default values in the services list that will be used for the scheduled tasks
          @{Name = "TaskPath"; Expression = { "N\A" } }, 
          @{Name = "Type"; Expression = { "Service" } }

          Write-Host "[$(Get-Date)] $(if ($Services -and -not $Services.Count) { 1 } else { $Services.Count }) $Lang found for $C"
          
          $Services | ForEach-Object { 
            Write-Host " - $($_.Name)$(If ($_.StartName) { " ($($_.StartName))" } )" 
          }

          $Results += $Services
        }
        else {
          Write-Host "[$(Get-Date)] No $Lang found for $C"
        }

        #update the output language
        $Lang = "scheduled tasks"
        Write-Host "[$(Get-Date)] Checking for $Lang"

        #We don't need the classname with the following cmdlet so remove it
        $Properties.Remove("ClassName")

        #Create an executable variable for the schtasks command so it's not created twice
        $SchTasks = { 
          $(if ($C -eq $env:COMPUTERNAME) { schtasks /query /V /FO CSV } else { schtasks /query /s $C /V /FO CSV }) | ConvertFrom-Csv | Where-Object { 
            $_.TaskName -notmatch "\\Microsoft\\" -and
            $_."Run As User" -ne "SYSTEM" -and 
            $_."Run As User" -ne "NETWORK SERVICE" -and 
            ($_."Run As User").Length -gt 1
          } | Select-Object @{Name = "ComputerName"; Expression = { $C } }, 
          @{Name = "Name"; Expression = { Split-Path $_.TaskName -Leaf } }, 
          @{Name = "StartName"; Expression = { $_."Run As User" } }, 
          @{Name = "StartMode"; Expression = { $_."Scheduled Task State" } }, 
          @{Name = "State"; Expression = { $_.Status } }, 
          @{Name = "TaskPath"; Expression = { $_.TaskName } }, 
          @{Name = "Type"; Expression = { "Task" } }
        }

        #If this OS is at least 2012 and 
        try {
          $Tasks = if ($OSVersion -ge 6.2) { 
            Get-ScheduledTask @Properties | Where-Object {
              $_.TaskPath -notmatch "Microsoft" -and 
              $_.Principal.UserId -ne "SYSTEM" -and 
              $_.Principal.UserId -ne "NETWORK SERVICE" -and 
              $_.Principal.UserId.Length -gt 1
            } | Select-Object @{Name = "ComputerName"; Expression = { $C } }, 
            @{Name = "Name"; Expression = { $_.TaskName } }, 
            @{Name = "StartName"; Expression = { $_.Principal.UserId } }, 
            @{Name = "StartMode"; Expression = { $_.Principal.RunLevel } }, 
            State, TaskPath, @{Name = "Type"; Expression = { "Task" } }
          }
          else { &$SchTasks }
        }
        catch { 
          try { 
            #Just in case there was an issue with CIM, lets try with the schtasks command
            if ($OSVersion -ge 6.2) { 
              &$SchTasks
            }
            else {
              throw "Error using the schtasks command"
            }
          }
          catch {
            Write-Host "[$(Get-Date)] $($_.Exception.Message)" -ForegroundColor Red
          }
        }

        # If tasks are collected output to screen, log and results for CSV output
        if ($Tasks) {
          #Use a name match for all tasks to exclude
          foreach ($E in $ExcludeTasks) { $Tasks = $Tasks | Where-Object { $_.Name -notmatch $E } }

          Write-Host "[$(Get-Date)] $(if ($Tasks -and -not $Tasks.Count) { 1 } else { $Tasks.Count }) $Lang found for $C"

          $Tasks | ForEach-Object { 
            Write-Host " - $($_.Name)$(If ($_.StartName) { " ($($_.StartName))" } )" 
          }

          $Results += $Tasks
        }
        else {
          Write-Host "[$(Get-Date)] No $Lang found for $C"
        }
      }
      else { Write-Host "[$(Get-Date)] Unable to connect to $C" }
    } #foreach($C in $ComputerName)

    # Export to CSV if enabled and ?{$_} to support older versions of PowerShell
    if ($CSVPath -and $Results) { 
      Write-Host "[$(Get-Date)] Exporting results to $CSVFile"
      $Results | Where-Object { $_ } | Export-Csv -Path $CSVFile -NoTypeInformation
    } 
    elseif ($CSVPath -and -not $Results) { 
      Write-Host "[$(Get-Date)] Either all results were filtered out or errors prevented collection. Check or use the Transcript to confirm."
    }

    if ($Transcript -and $LogFile) { Stop-Transcript }
  }
}

#Bind the parameters to our function
Get-ServiceAccounts @PSBoundParameters