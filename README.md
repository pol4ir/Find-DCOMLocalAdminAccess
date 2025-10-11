# Find-DCOMLocalAdminAccess
Enumerate accessible machines across a domain or local network via DCOM-based techniques

## Usage
<img src="https://github.com/pol4ir/Find-DCOMLocalAdminAccess/blob/main/test.gif">

```
Find-DCOMLocalAdminAccess -lhost 192.168.56.30
```
```
Find-DCOMLocalAdminAccess -lhost 192.168.56.30 -lport 6789 -threads 5 -Computername dc.contoso.local -timeout 10000
```

### RunAs
The script runs under the current user session. If you're in an interactive shell and need to execute it under a different security context, you can use Runas.
```
runas /user:contoso.local\user1 /netonly powershell
```
If you're working in a non-interactive shell, you can use <a href="https://github.com/antonioCoco/RunasCs">Invoke-RunasCs</a>

```
Invoke-RunasCs -Domain contoso.local -Username user1 -Password dfgV?DS7-8 -Command 'powershell . "C:\Find-DCOMLocalAdminAccess.ps1";Find-DCOMLocalAdminAccess -lhost "192.168.56.30"' -logontype 9
```
