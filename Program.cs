using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XmlWriter
{
    internal static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                Console.OutputEncoding = System.Text.Encoding.UTF8;
                if (args.Length < 3)
                {
                    Console.WriteLine("Usage: XmlWriter.exe <ExcelPath> <OutputDir> <TemplateFilePath>");
                    return;
                }
                try
                {
                    XmlWriter.Utility.CommandRunner.Run(args[0], args[1], args[2]);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    Console.WriteLine(ex.StackTrace);
                }
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
