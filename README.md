# TQM User export

## Setup

1. Create a `envs.ps1`
1. Fill it with:
    ```PowerShell
    $papertrailConfig = @{
        Server = "<papertrailserver>"
        Port = <portnumber>
        HostName = "<hostname>"
        Level = "ERROR" # papertrail logging will be activated from this log level and up
    }
    ```

## Usage

1. From a PowerShell prompt
    1. .\Start-UserExport.ps1
