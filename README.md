# MiscTools

## CsExec
Command Exec / Lateral movement via PsExec-like functionality.  Must be running in the context of a privileged user.

### Usage
```
CsExec.exe <targetMachine> <serviceName> <serviceDisplayName> <binPath>
```

Also see [TikiService](https://rastamouse.me/2019/08/tikiservice/).

## CsWMI
Command Exec / Lateral Movement via WMI. Must be running in the context of a privileged user.

Current methods: `ProcessCallCreate`.

### Usage
```
CsWMI.exe <targetMachine> <command> <method>
```

Also see [The Return of Aggressor](https://rastamouse.me/2019/06/the-return-of-aggressor/)

## CsDCOM
Command Exec / Lateral Movement via DCOM. Must be running in the context of a privileged user.

Current Methods: `MMC20.Application`, `ShellWindows`, `ShellBrowserWindow`, `ExcelDDE`.

### Usage
```
CsDCOM.exe <targetMachine> <binary> <arg> <method>
```

## CsEnv
Add user/machine/process environment variables.

### Usage:
```
CsEnv.exe <variableName> <value> <target>
```

## Credits
Most code blatently stolen and adapted from:
- [Invoke-PsExec](https://github.com/EmpireProject/Empire/blob/master/data/module_source/lateral_movement/Invoke-PsExec.ps1) by [harmj0y](https://twitter.com/harmj0y)
- [SharpWMI](https://github.com/GhostPack/SharpWMI) by [harmj0y](https://twitter.com/harmj0y)
- [SharpCOM](https://github.com/rvrsh3ll/SharpCOM) by [rvrsh3ll](https://twitter.com/424f424f)