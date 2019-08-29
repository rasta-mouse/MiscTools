# MiscTools

## CsExec
(Some) PsExec functionality in C#.  Must be running in the context of a privileged user.

### Usage
```
CsExec.exe <targetMachine> <serviceName> <serviceDisplayName> <binPath>
```

Also see [TikiService](https://rastamouse.me/2019/08/tikiservice/).

## WMI
Command Exec / Lateral Movement via WMI. Must be running in the context of a privileged user.
Current methods: `ProcessCallCreate`.

### Usage
```
WMI.exe <targetMachine> <command> <method>
```

Also see [The Return of Aggressor](https://rastamouse.me/2019/06/the-return-of-aggressor/)