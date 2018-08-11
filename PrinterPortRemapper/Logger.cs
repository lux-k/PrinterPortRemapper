using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PrinterPortRemapper
{
    static class Logger
    {
        private static StreamWriter sw = null;
        private static void Init()
        {
            sw = new StreamWriter(Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)) + "\\printer-remapper.txt", true);
        }

        public static void Log(string s)
        {
            if (sw == null)
                Init();

            Console.WriteLine(s);
            sw.WriteLine(s);
        }

        public static void Cleanup() {
            if (sw != null)
                sw.Close();

        }
    }
}
