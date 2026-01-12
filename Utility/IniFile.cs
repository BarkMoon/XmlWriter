using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace XmlWriter
{
    public class IniFile
    {
        public string Path;

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        // コンストラクタ（実行ファイルと同じ場所にSettings.iniを作成・参照します）
        public IniFile(string iniPath = null)
        {
            Path = new FileInfo(iniPath ?? System.IO.Path.Combine(Application.StartupPath, "Settings.ini")).FullName;
        }

        public void Write(string key, string value, string section = "Settings")
        {
            WritePrivateProfileString(section, key, value, Path);
        }

        public string Read(string key, string section = "Settings")
        {
            var retVal = new StringBuilder(255);
            GetPrivateProfileString(section, key, "", retVal, 255, Path);
            return retVal.ToString();
        }
    }
}