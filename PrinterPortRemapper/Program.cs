using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.ServiceProcess;
using System.Net;

// needs to run as administrator because of the registry changes
// and to restart print spooler

namespace PrinterPortRemapper
{
    class Program
    {
        const String BASE = "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Print\\Printers";
        delegate bool PortProcessor(String s);
        static Dictionary<String, PortProcessor> PortMap;

        static void Main(string[] args)
        {
            PortMap = LoadPortMap();
            bool changes = false;
            Console.WriteLine("Enumerating printers and processing known ports\n---");
            foreach (String s in GetPrinters())
            {
                changes |= EvaluatePrinter(s);
            }

            if (changes)
                RestartSpooler();

            Console.ReadLine();
        }

        static Dictionary<String, PortProcessor> LoadPortMap()
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

        static bool EvaluatePrinter(String s)
        {
            RegistryKey cc = Registry.LocalMachine.OpenSubKey(BASE + "\\" + s);
            Console.WriteLine("Name: " + cc.GetValue("Name"));
            Console.WriteLine("Port Name: " + cc.GetValue("Port"));
            bool res = EvaluatePrinterPort(cc.GetValue("Port").ToString());
            Console.WriteLine("----");
            return res;
        }


        static bool WSDPort(String s)
        {
            Console.WriteLine("Processing " + s + " as a WSD port.");
            String guid;
            bool res = false;
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
                                    res = true;
                                }
                            }
                        }
                    }
                }
            }
            
            return res;
        }

        static bool StandardPort(String s)
        {
            Console.WriteLine("Processing " + s + " as a Standard TCP/IP port.");

            RegistryKey r = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Print\\Monitors\\Standard TCP/IP Port\\Ports\\" + s, true);

            if (r != null && r.GetValue("HostName") != null)
            {
                String val = r.GetValue("HostName").ToString();
                if (System.Net.IPAddress.TryParse(val, out IPAddress ip))
                {
                    String host = IPtoHostname(ip);
                    if (host != null)
                    {
                        r.SetValue("HostName", host);
                        Console.WriteLine("Changing hostname from " + val + " to " + host);
                        return true;
                    }

                }
            }
            return false;
        }

        static bool EvaluatePrinterPort(String s)
        {
            bool res = false;
            if (s.StartsWith("WSD-"))
            {
                res |= WSDPort(s);
            } else {
                // invoked the mapped processor
                if (PortMap.ContainsKey(s))
                {
                    res |= PortMap[s](s);
                }
                 else 
                    Console.WriteLine("Don't know how to deal with port " + s);
            }
            return res;
        }

        static String IPtoHostname(IPAddress ip)
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
    }
}
