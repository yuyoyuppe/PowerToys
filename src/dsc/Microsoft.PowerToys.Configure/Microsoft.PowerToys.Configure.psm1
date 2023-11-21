#region enums
enum PowerToysConfigureScope {
    Machine
    User
}

enum PowerToysConfigureEnsure {
    Absent
    Present
}
#endregion enums

#region DscResources
[DscResource()]
class PowerToysConfigure {
    [DscProperty()] [PowerToysConfigureEnsure]
    $Ensure = [PowerToysConfigureEnsure]::Present

    [DscProperty(Key)]
    [string]$BackupFile = "C:\\settings_133445405941718563.ptb"

    [PowerToysConfigure] Get() {
        $CurrentState = [PowerToysConfigure]::new()
        return @{
            BackupFile = $this.BackupFile
        }
    }

    [bool] Test() {
        return $true
    }

    [void] Set() {
        $this.Get()

        $TempFilePath = Join-Path -Path $env:TEMP -ChildPath "TestConfigure.txt"
        Set-Content -Path "$TempFilePath" -Value $this.BackupFile -Force
    }
}
#endregion DscResources

