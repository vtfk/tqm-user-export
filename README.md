# TQM User export

## Setup

1. Create a `envs.ps1`
1. Fill it with:
    ```PowerShell
    $papertrailConfig = @{ # Optional
        Server = "<papertrailserver>"
        Port = <portnumber>
        HostName = "<hostname>"
        Level = "ERROR" # papertrail logging will be activated from this log level and up
    }

    $sftp = @{
        HostName = "sftp.server.no"
        UserName = "username"
        Password = "password"
        SshHostKeyFingerprint = "ssh-rsa 4096 ssh-fingerprint"
    }

    $tqmConfig = @{
        Firmakode = "firmakoden din"
        TeamsWebhook = "https://vtfk.webhook.office.com/webhookb2/blablabla" # Sends teams webhook on logger error level
    }
    ```

## Usage

1. From a PowerShell prompt
    1. `.\Start-UserExport.ps1`

## Start-ScriptFile.ps1

This will file be launched by a scheduled task. ~~Only when there's one or more file(s) at the `$folder` location, `Start-UserExport.ps1` will be launched.~~ Whenever you set it to run (TQM likes around midnight)
