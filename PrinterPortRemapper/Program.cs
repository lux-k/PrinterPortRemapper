using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.ServiceProcess;
using System.Net;

// needs to run as administrator because of the registry changes
// and to restart print spooler

namespace PrinterPortRemapper
{
    class PortStatus
    {
        private bool _c;
        private string _n = null;
        public PortStatus()
        {
            Changed = false;
        }
        public PortStatus(bool c)
        {
            Changed = c;
        }
        public PortStatus(bool c, string n )
        {
            Changed = c;
            NewName = n;
        }


        public bool Changed
        {
            get
            {
                return _c;
            }
            set
            {
                _c = value;
            }
        }

        public string NewName
        {
            get
            {
                return _n;
            }
            set
            {
                _n = value;
            }
        }
    }

    class Program
    {
        const string BASE = "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Print\\Printers";
        delegate PortStatus PortProcessor(string s);
        static Dictionary<string, PortProcessor> PortMap;

        static void Main(string[] args)
        {
            PortMap = LoadPortMap();
            bool changes = false;
            Console.WriteLine("Enumerating printers and processing known ports\n---");
            foreach (string s in GetPrinters())
            {
                changes |= EvaluatePrinter(s);
            }

            if (changes)
                RestartSpooler();

            Console.ReadLine();
        }

        static Dictionary<string, PortProcessor> LoadPortMap()
        {
            Console.WriteLine("---\nEnumerating known ports");
            Dictionary<String, PortProcessor> PortMap = new Dictionary<string, PortProcessor>();

            RegistryKey r = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Print\\Monitors");

            // standard tcp port
            {
                RegistryKey kid = r.OpenSubKey("Standard TCP/IP Port");
                if (kid != null)
                {
                    kid = kid.OpenSubKey("Ports");
                    if (kid != null)
                    {
                        foreach (String p in kid.GetSubKeyNames())
                        {
                            Console.WriteLine("Found Standard TCP/IP port: " + p);
                            PortMap[p] = new PortProcessor(StandardPort);
                        }
                    }
                } 
                
            }
            Console.WriteLine("---");
            return PortMap;
        }

        static String[] GetPrinters( )
        {
            RegistryKey cc = Registry.LocalMachine.OpenSubKey(BASE);
            String[] names = cc.GetSubKeyNames();
            return names;
        }

        static void RestartSpooler()
        {
            Console.WriteLine("Restarting the print spooler");
            ServiceController service = new ServiceController("Spooler");
            TimeSpan timeout = TimeSpan.FromMilliseconds(30000);

            service.Stop();
            service.WaitForStatus(ServiceControllerStatus.Stopped, timeout);
            service.Start();
            Console.WriteLine("Restart of spooler is complete.");
        }

        static bool EvaluatePrinter(string s)
        {
            RegistryKey cc = Registry.LocalMachine.OpenSubKey(BASE + "\\" + s, true);
            Console.WriteLine("Name: " + cc.GetValue("Name"));
            string port = cc.GetValue("Port").ToString();
            Console.WriteLine("Port Name: " + port);
            PortStatus res = EvaluatePrinterPort(port);
            if (res != null && res.Changed && res.NewName != null)
            {
                Console.WriteLine("The printer port name has changed. Changing port name from " + port + " to " + res.NewName);
                cc.SetValue("Port", res.NewName);
            }

            Console.WriteLine("----");
            return res != null && res.Changed;
        }

        static PortStatus WSDPort(string s)
        {
            Console.WriteLine("Processing " + s + " as a WSD port.");
            String guid;
            PortStatus res = new PortStatus();
            RegistryKey r = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\SWD\\PRINTENUM\\" + s);
            if (r != null && r.GetValue("ContainerID") != null) {
                guid = r.GetValue("ContainerID").ToString();
                guid = guid.Substring(1, guid.Length - 2);
                Console.WriteLine("ContainerID is " + guid);
                r = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\SWD\\DAFWSDProvider");
                if (r != null)
                {
                    foreach (String n in r.GetSubKeyNames())
                    {
                        if (n.StartsWith("urn:uuid:"+guid))
                        {
                            Console.WriteLine(n + " matches the WSD set.");
                            RegistryKey kid = r.OpenSubKey(n, true);
                            String loc = kid.GetValue("LocationInformation").ToString();
                            Uri u = new Uri(loc);
                            if (System.Net.IPAddress.TryParse(u.Host, out IPAddress ip)) { 
                                String host = IPtoHostname(ip);
                                if (host != null)
                                {
                                    String newloc = loc.Replace(ip.ToString(), host);
                                    Console.WriteLine("Changing URL from " + loc + " to " + newloc);
                                    kid.SetValue("LocationInformation", newloc);
                                    res.Changed = true;
                                }
                            }
                        }
                    }
                }
            }
            
            return res;
        }

        static PortStatus StandardPort(string s)
        {
            Console.WriteLine("Processing " + s + " as a Standard TCP/IP port.");

            RegistryKey r = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Print\\Monitors\\Standard TCP/IP Port\\Ports\\" + s, true);

            if (r != null && r.GetValue("HostName") != null)
            {
                string val = r.GetValue("HostName").ToString();
                if (System.Net.IPAddress.TryParse(val, out IPAddress ip))
                {
                    string host = IPtoHostname(ip);
                    if (host != null)
                    {
                        r.SetValue("HostName", host);
                        Console.WriteLine("Changing hostname from " + val + " to " + host);
                        string nn = GetHostname(host);
                        r = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Print\\Monitors\\Standard TCP/IP Port\\Ports\\",true);
                        RenameSubKey(r, s, nn);
                        return new PortStatus(true, nn);
                    }

                }
            }
            return new PortStatus(false);
        }

        static PortStatus EvaluatePrinterPort(string s)
        {
            if (s.StartsWith("WSD-"))
            {
                return WSDPort(s);
            } else {
                // invoked the mapped processor
                if (PortMap.ContainsKey(s))
                {
                   return PortMap[s](s);
                }
                 else 
                    Console.WriteLine("Don't know how to deal with port " + s);
            }

            return null;
        } 

        static string IPtoHostname(IPAddress ip)
        {
            IPHostEntry hostInfo = null;
            try
            {
                hostInfo = Dns.GetHostEntry(ip);
            }
            catch { }

            if (hostInfo != null && hostInfo.HostName != null)
            {
                return hostInfo.HostName;
            } else
            {
                Console.WriteLine("Unable to find a hostname for " + ip);
            }
            return null;
        }

        static string GetHostname(string s )
        {
            return s.Substring(0, s.IndexOf('.') - 1);
        }

        // from https://www.codeproject.com/Articles/16343/Copy-and-Rename-Registry-Keys
        static bool RenameSubKey(RegistryKey parentKey, string subKeyName, string newSubKeyName)
        {
            CopyKey(parentKey, subKeyName, newSubKeyName);
            parentKey.DeleteSubKeyTree(subKeyName);
            return true;
        }

        /// <summary>
        /// Copy a registry key.  The parentKey must be writeable.
        /// </summary>
        /// <param name="parentKey"></param>
        /// <param name="keyNameToCopy"></param>
        /// <param name="newKeyName"></param>
        /// <returns></returns>
        static bool CopyKey(RegistryKey parentKey,
            string keyNameToCopy, string newKeyName)
        {
            //Create new key
            RegistryKey destinationKey = parentKey.CreateSubKey(newKeyName);

            //Open the sourceKey we are copying from
            RegistryKey sourceKey = parentKey.OpenSubKey(keyNameToCopy);

            RecurseCopyKey(sourceKey, destinationKey);

            return true;
        }

        static void RecurseCopyKey(RegistryKey sourceKey, RegistryKey destinationKey)
        {
            //copy all the values
            foreach (string valueName in sourceKey.GetValueNames())
            {
                object objValue = sourceKey.GetValue(valueName);
                RegistryValueKind valKind = sourceKey.GetValueKind(valueName);
                destinationKey.SetValue(valueName, objValue, valKind);
            }

            //For Each subKey 
            //Create a new subKey in destinationKey 
            //Call myself 
            foreach (string sourceSubKeyName in sourceKey.GetSubKeyNames())
            {
                RegistryKey sourceSubKey = sourceKey.OpenSubKey(sourceSubKeyName);
                RegistryKey destSubKey = destinationKey.CreateSubKey(sourceSubKeyName);
                RecurseCopyKey(sourceSubKey, destSubKey);
            }
        }
    }
}
