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
