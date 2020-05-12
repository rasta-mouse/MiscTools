# MiscTools

## CsExec
Command Exec / Lateral movement via PsExec-like functionality.  Must be running in the context of a privileged user.

```
CsExec.exe <targetMachine> <serviceName> <serviceDisplayName> <binPath>
```

Also see [TikiService](https://rastamouse.me/2019/08/tikiservice/).

## CsPosh
Command Exec / Lateral Movement via PowerShell. Creates a PowerShell runspace on a remote target.
Must be running in the context of a privileged user.

```
Usage:
  -t, --target=VALUE         Target machine
  -c, --code=VALUE           Code to execute
  -e, --encoded              Indicates that provided code is base64 encoded
  -o, --outstring            Append Out-String to code
  -r, --redirect             Redirect stderr to stdout
  -d, --domain=VALUE         Domain for alternate credentials
  -u, --username=VALUE       Username for alternate credentials
  -p, --password=VALUE       Password for alternate credentials
  -h, -?, --help             Show Help
```

## CsWMI
Command Exec / Lateral Movement via WMI. Must be running in the context of a privileged user.

Current methods: `ProcessCallCreate`.

```
CsWMI.exe <targetMachine> <command> <method>
```

Also see [The Return of Aggressor](https://rastamouse.me/2019/06/the-return-of-aggressor/)

## CsDCOM
Command Exec / Lateral Movement via DCOM. Must be running in the context of a privileged user.

Current Methods: `MMC20.Application`, `ShellWindows`, `ShellBrowserWindow`, `ExcelDDE`.

```
Usage:
  -t, --target=VALUE         Target Machine
  -b, --binary=VALUE         Binary: powershell.exe
  -a, --args=VALUE           Arguments: -enc <blah>
  -m, --method=VALUE         Method: MMC20Application, ShellWindows,
                               ShellBrowserWindow, ExcelDDE
  -h, -?, --help             Show Help
```

## CsEnv
Add user/machine/process environment variables.

```
CsEnv.exe <variableName> <value> <target>
```

## Credits
Most code blatently stolen and adapted from:
- [Invoke-PsExec](https://github.com/EmpireProject/Empire/blob/master/data/module_source/lateral_movement/Invoke-PsExec.ps1) by [harmj0y](https://twitter.com/harmj0y)
- [SharpWMI](https://github.com/GhostPack/SharpWMI) by [harmj0y](https://twitter.com/harmj0y)
- [SharpCOM](https://github.com/rvrsh3ll/SharpCOM) by [rvrsh3ll](https://twitter.com/424f424f)