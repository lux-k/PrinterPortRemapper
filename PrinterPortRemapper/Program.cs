using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.ServiceProcess;
using System.Net;
using System.Collections;
using System.Linq;

// needs to run as administrator because of the registry changes
// and to restart print spooler

namespace PrinterPortRemapper
{
    public enum ProgramMode { ToHost, ToIP, ListPorts };
    public enum SNMPMode { ForceOff, ForceOn, Ignore };

    class NetworkRange
    {
        private IPAddress start;
        private IPAddress end;

        public NetworkRange()
        {
            start = IPAddress.Parse("0.0.0.0");
            end = IPAddress.Parse("255.255.255.255");
        }

        public NetworkRange(string s)
        {
            if (s.Contains("-"))
            {
                string[] arr = s.Split('-');
                start = IPAddress.Parse(arr[0]);
                end = IPAddress.Parse(arr[1]);
            }
            else
            {
                start = IPAddress.Parse(s);
                end = IPAddress.Parse(s);
            }
        }

        public bool InRange(IPAddress a)
        {
            UInt32 ipStart = BitConverter.ToUInt32(start.GetAddressBytes().Reverse().ToArray(), 0);
            UInt32 ipEnd = BitConverter.ToUInt32(end.GetAddressBytes().Reverse().ToArray(), 0);
            UInt32 ip = BitConverter.ToUInt32(a.GetAddressBytes().Reverse().ToArray(), 0);
            return ipStart <= ip && ip <= ipEnd;
        }

        public bool InRange(string s)
        {
            return InRange(IPAddress.Parse(s));
        }

        public override string ToString()
        {
            return start.ToString() + " - " + end.ToString();
        }
    }

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
        public PortStatus(bool c, string n)
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
        static ProgramMode Mode = ProgramMode.ListPorts;
        static SNMPMode SNMP = SNMPMode.Ignore;
        static ArrayList _networkRanges = new ArrayList();
        static Dictionary<string, string> ManualDictionary = new Dictionary<string, string>();

        static void Main(string[] args)
        {
            Logger.Log("----");
            Logger.Log("Program log start: " + DateTime.Now);
            Logger.Log("Machine name: " + Environment.MachineName);
            Logger.Log("User name: " + Environment.UserName);


            foreach (string s in args)
            {
                if (s.Equals("toip"))
                    Mode = ProgramMode.ToIP;
                else if (s.Equals("tohost"))
                    Mode = ProgramMode.ToHost;
                else if (s.Equals("disablesnmp"))
                    SNMP = SNMPMode.ForceOff;
                else if (s.Equals("enablesnmp"))
                    SNMP = SNMPMode.ForceOn;
                else if (s.Equals("ignoresnmp"))
                    SNMP = SNMPMode.Ignore;
                else if (s.StartsWith("range:"))
                {
                    foreach (string range in s.Substring(6).Split(','))
                    {
                        _networkRanges.Add(new NetworkRange(range));
                    }
                }
                else if (s.StartsWith("rename:"))
                {
                    foreach (string range in s.Substring(7).Split(','))
                    {
                        string[] arr = range.Split('=');
                        ManualDictionary.Add(arr[0], arr[1]);
                    }
                }
            }

            if (_networkRanges.Count == 0)
                _networkRanges.Add(new NetworkRange());

            Logger.Log("Translating to IP: " + (Mode == ProgramMode.ToIP).ToString());
            Logger.Log("Translating to Hostname: " + (Mode == ProgramMode.ToHost).ToString());
            Logger.Log("SNMP flag: " + (SNMP.ToString()));
            Logger.Log("Valid ranges: " + String.Join(", ", _networkRanges.ToArray().Select( el => el.ToString())));

            PortMap = LoadPortMap();
            bool changes = false;
            Logger.Log("Enumerating and processing printers.\n---");
            foreach (string s in GetPrinters())
            {
                changes |= EvaluatePrinter(s);
            }

            if (changes)
                RestartSpooler();

            Logger.Cleanup();
        }

        static bool InAllowedRanges(IPAddress i)
        {
            foreach (NetworkRange range in _networkRanges)
            {
                bool res = range.InRange(i);
                if (res)
                {
                    Logger.Log("Process " + i.ToString() + " : " + range.ToString());
                    return true;
                }

            }

            Logger.Log(i.ToString() + " not found on allowed network ranges.");
            return false;
        }

        static bool InAllowedRanges(string s)
        {
            IPAddress i = null;
            IPAddress.TryParse(s, out i);
            if (i == null)
            {
                return false;
            }
            else
            {
                return InAllowedRanges(i);
            }

        }

        static Dictionary<string, PortProcessor> LoadPortMap()
        {
            Logger.Log("---\nEnumerating known port types");
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
                            Logger.Log("Found Standard TCP/IP port: " + p);
                            PortMap[p] = new PortProcessor(StandardPort);
                        }
                    }
                }

            }
            Logger.Log("---");
            return PortMap;
        }

        static String[] GetPrinters()
        {
            RegistryKey cc = Registry.LocalMachine.OpenSubKey(BASE);
            String[] names = cc.GetSubKeyNames();
            return names;
        }

        static void RestartSpooler()
        {
            Logger.Log("Restarting the print spooler");
            ServiceController service = new ServiceController("Spooler");
            TimeSpan timeout = TimeSpan.FromMilliseconds(30000);

            service.Stop();
            service.WaitForStatus(ServiceControllerStatus.Stopped, timeout);
            service.Start();
            Logger.Log("Restart of spooler is complete.");
        }

        static bool EvaluatePrinter(string s)
        {
            RegistryKey cc = Registry.LocalMachine.OpenSubKey(BASE + "\\" + s, true);
            Logger.Log("Name: " + cc.GetValue("Name"));
            string port = cc.GetValue("Port").ToString();
            Logger.Log("Port Name: " + port);
            PortStatus res = EvaluatePrinterPort(port);
            if (res != null && res.Changed && res.NewName != null)
            {
                Logger.Log("The printer port name has changed. Changing port name from " + port + " to " + res.NewName);
                cc.SetValue("Port", res.NewName);
            }

            Logger.Log("----");
            return res != null && res.Changed;
        }

        static PortStatus WSDPort(string s)
        {
            Logger.Log("Processing " + s + " as a WSD port.");
            String guid;
            PortStatus res = new PortStatus();
            RegistryKey reg = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\SWD\\PRINTENUM\\" + s, true);
            String host = null;
            if (reg != null && reg.GetValue("ContainerID") != null)
            {
                guid = reg.GetValue("ContainerID").ToString();
                guid = guid.Substring(1, guid.Length - 2);
                Logger.Log("ContainerID is " + guid);
                RegistryKey r = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\SWD\\DAFWSDProvider", true);
                if (r != null)
                {
                    foreach (String n in r.GetSubKeyNames())
                    {
                        if (n.StartsWith("urn:uuid:" + guid) || n.StartsWith("uuid:" + guid))
                        {
                            Logger.Log(n + " matches the WSD set.");
                            RegistryKey kid = r.OpenSubKey(n, true);
                            String loc = kid.GetValue("LocationInformation").ToString();
                            Logger.Log("WSD URI: " + loc);
                            Uri u = new Uri(loc);
                            System.Net.IPAddress.TryParse(u.Host, out IPAddress ip);

                            if (Mode == ProgramMode.ToHost && ip != null && InAllowedRanges(ip))
                                host = IPtoHostname(ip);
                            else if (Mode == ProgramMode.ToIP && ip == null)
                            {
                                host = HostnameToIP(u.Host);
                                if (!InAllowedRanges(host))
                                    host = null;
                            }

                            if (host != null)
                            {
                                String newloc = ReplaceIgnoreCase(loc, u.Host, host);
                                Logger.Log("Changing URL from " + loc + " to " + newloc);
                                kid.SetValue("LocationInformation", newloc);
                                res.Changed = true;
                            }
                        }
                    }
                }

                if (res.Changed)
                {
                    if (reg != null && reg.GetValue("LocationInformation") != null)
                    {
                        String loc = reg.GetValue("LocationInformation").ToString();
                        Logger.Log("WSD URI: " + loc);
                        Uri u = new Uri(loc);
                        String newloc = ReplaceIgnoreCase(loc, u.Host, host);
                        if (!newloc.Equals(loc))
                        {
                            Logger.Log("Changing URL from " + loc + " to " + newloc);
                            reg.SetValue("LocationInformation", newloc);
                        }
                    }
                }
            }
            return res;
        }

        static string ReplaceIgnoreCase(string s, string word, string replace)
        {
            int index = s.IndexOf(word, StringComparison.OrdinalIgnoreCase);

            // perform the replace on the matched word
            if (index >= 0)
            {
                s = s.Remove(index, word.Length);
                s = s.Insert(index, replace);
            }
            return s;
        }

        static PortStatus StandardPort(string s)
        {
            Logger.Log("Processing " + s + " as a Standard TCP/IP port.");

            RegistryKey r = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Print\\Monitors\\Standard TCP/IP Port\\Ports\\" + s, true);


            if (SNMP != SNMPMode.Ignore)
            {
                string myKey = "SNMP Enabled";
                if (r.GetValue(myKey) != null)
                {
                    string snmpVal = r.GetValue(myKey).ToString();
                    if (snmpVal.Equals("1") && SNMP == SNMPMode.ForceOff)
                    {
                        Logger.Log("Turning OFF SNMP...");
                        r.SetValue(myKey, 0);
                    }
                    else if (snmpVal.Equals("0") && SNMP == SNMPMode.ForceOn)
                    {
                        Logger.Log("Turning ON SNMP...");
                        r.SetValue(myKey, 1);
                    }
                }
            }

            if (r != null && r.GetValue("HostName") != null)
            {
                string val = r.GetValue("HostName").ToString();
                System.Net.IPAddress.TryParse(val, out IPAddress ip);
                if (Mode == ProgramMode.ToHost && ip != null && InAllowedRanges(ip))
                {
                    string host = IPtoHostname(ip);
                    if (host != null)
                    {
                        r.SetValue("HostName", host);
                        Logger.Log("Changing hostname from " + val + " to " + host);
                        string nn = GetHostname(host);
                        if (!val.Equals(host))
                        {
                            r = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Print\\Monitors\\Standard TCP/IP Port\\Ports", true);
                            nn = "PMACS_" + nn;
                            RenameSubKey(r, s, nn);
                        }
                        return new PortStatus(true, nn);
                    }
                }
                else if (Mode == ProgramMode.ToIP && ip == null)
                {
                    string host = HostnameToIP(val);
                    if (host != null && InAllowedRanges(host))
                    {
                        r.SetValue("HostName", host);
                        Logger.Log("Changing hostname from " + val + " to " + host);
                        string nn = host.Replace('.', '_');
                        if (!val.Equals(host))
                        {
                            r = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Print\\Monitors\\Standard TCP/IP Port\\Ports", true);
                            nn = "PMACS_" + nn;
                            RenameSubKey(r, s, nn);SY
                        }
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
            }
            else
            {
                // invoked the mapped processor
                if (PortMap.ContainsKey(s))
                {
                    return PortMap[s](s);
                }
            }

            return null;
        }

        static string HostnameToIP(string host)
        {
            if (ManualDictionary.ContainsKey(host))
            {
                Logger.Log("Using manual entry of " + host + " --> " + ManualDictionary[host]);
                return ManualDictionary[host];
            }

            string ip = null;
            try
            {
                ip = Dns.GetHostAddresses(host)[0].ToString();
            }
            catch { }

            if (ip != null)
                return ip;
            else
            {
                Logger.Log("Unable to find IP for host " + host);
                return null;
            }
        }

        static string IPtoHostname(IPAddress ip)
        {
            if (ManualDictionary.ContainsKey(ip.ToString()))
            {
                Logger.Log("Using manual entry of " + ip.ToString() + " --> " + ManualDictionary[ip.ToString()]);
                return ManualDictionary[ip.ToString()];
            }

            IPHostEntry hostInfo = null;
            try
            {
                hostInfo = Dns.GetHostEntry(ip);
            }
            catch { }

            if (hostInfo != null && hostInfo.HostName != null)
            {
                return hostInfo.HostName;
            }
            else
            {
                Logger.Log("Unable to find a hostname for " + ip);
            }
            return null;
        }

        static string GetHostname(string s)
        {
            return s.Substring(0, s.IndexOf('.'));
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
