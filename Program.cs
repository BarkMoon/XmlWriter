using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XmlWriter
{
    internal static class Program
    {
        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeConsole();

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                try { Console.OutputEncoding = System.Text.Encoding.UTF8; } catch { }
                try
                {
                    // 簡易引数パーサー
                    var options = new XmlWriter.Utility.CommandRunner.RunOptions();
                    
                    // 第1引数はExcelパス（必須）とみなす
                    options.ExcelPath = args[0];

                    for (int i = 1; i < args.Length; i++)
                    {
                        string arg = args[i];
                        if (arg.StartsWith("-"))
                        {
                            string val = (i + 1 < args.Length) ? args[i + 1] : null;
                            
                            switch (arg.ToLower())
                            {
                                case "-mode":
                                case "-m":
                                    if (string.IsNullOrEmpty(val)) {
                                        Console.WriteLine("Mode argument missing value.");
                                        return;
                                    }
                                    if (Enum.TryParse(val, true, out XmlWriter.Utility.CommandRunner.ExecutionMode mode))
                                    {
                                        options.Mode = mode;
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Invalid mode: {val}");
                                        return;
                                    }
                                    i++; // consume value
                                    break;
                                case "-output":
                                case "-o":
                                    options.OutputDir = val;
                                    i++;
                                    break;
                                case "-template":
                                case "-t":
                                    options.TemplateFilePath = val;
                                    i++;
                                    break;
                                case "-target":
                                case "-table":
                                    options.TargetTableName = val;
                                    i++;
                                    break;
                                default:
                                    Console.WriteLine($"Unknown argument: {arg}");
                                    break;
                            }
                        }
                    }

                    if (string.IsNullOrEmpty(options.ExcelPath))
                    {
                         Console.WriteLine("Usage: XmlWriter.exe <ExcelPath> [options]");
                         return;
                    }

                    XmlWriter.Utility.CommandRunner.Run(options);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    Console.WriteLine(ex.StackTrace);
                }
                return;
            }

            // GUIモード: コンソールウィンドウを解放（隠す）
            FreeConsole();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
