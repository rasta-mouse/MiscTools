using System;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using System.Threading;

using static CsExec.NativeMethods;


namespace CsExec
{

    static class SimplePathValidator
    {
        static public bool IsValidPath(string path)
        {
            Regex r = new Regex(@"^(([a-zA-Z]:|\\\\\w[ \w\.]*)(\\\w*[ \w\.]*|\\+%[ \w\.]+%+)+|%[ \w\.]+%(\\\w[ \w\.]*|\\%[ \w\.]+%+)*)");
            return r.IsMatch(path);
        }
    }

    public class ServiceExecutor
    {
        private IntPtr _svcman;
        private IntPtr _svc;
        private string _target;
        private string _serviceName;
        private string _serviceDisplayName;
        private string _binPath;
        SERVICE_STATUS _serviceStatus;


        public SERVICE_STATUS ServiceStatus
        {
            get => this._serviceStatus;
            set
            {
                this._serviceStatus = value;
            }
        }

        public IntPtr SvcManager
        {
            get => this._svcman;
            set
            {
                if ((value != null))
                {
                    this._svcman = value;
                }
            }
        }

        public IntPtr Service
        {
            get => this._svc;
            set
            {
                if ((value != null) )
                {
                    this._svc = value;
                }
            }
        }

        public string ServiceName
        {
            get => this._serviceName;
            set
            {
                this._serviceName = value;
            }
        }

        public string ServiceDisplayName
        {
            get => this._serviceDisplayName;
            set
            {
                this._serviceDisplayName = value;
            }
        }

        public string ServiceExecutable
        {
            get => this._binPath;
            set
            {
                if (SimplePathValidator.IsValidPath(value))
                {
                    this._binPath = value;
                }
            }
        }

        public string Target
        {
            get => this._target;
            set
            {
                this._target = value;
            }
        }


        public void CallSvcManager()
        {
            if ((this.Target != null) && (this.Target != ""))
            {
                this.SvcManager = OpenSCManager(this.Target, null, SCM_ACCESS.SC_MANAGER_ALL_ACCESS);

            }
        }

        public bool Create()
        {

            if (this.SvcManager == IntPtr.Zero)
            {
                CallSvcManager();
            }
            this.Service = CreateService(this.SvcManager, this.ServiceName, this.ServiceDisplayName, SERVICE_ACCESS.SERVICE_ALL_ACCESS, SERVICE_TYPE.SERVICE_WIN32_OWN_PROCESS, SERVICE_START.SERVICE_DEMAND_START, SERVICE_ERROR.SERVICE_ERROR_NORMAL, this.ServiceExecutable, null, null, null, null, null);
            Console.WriteLine($"[+] Service Pointer: {this.Service.ToString()}");
            return this.Service != IntPtr.Zero;
        }

        public bool Start()
        {
            bool checkpoint = false;
            if (this.SvcManager == IntPtr.Zero)
            {
                CallSvcManager();
            }
            if (this.Service == IntPtr.Zero)
            {
                try
                {
                    GetServiceHandle();
                }
                catch {
                    Console.WriteLine($"[-] Unknown Service: {this.ServiceName}");
                    Console.WriteLine($"[*] Creating Service: {this.ServiceName}");
                    checkpoint = this.Create();
                }
            }
            if (checkpoint) { 
                StartService(this.Service, 0, null);
                this.Wait();
            }
            return IsRunning();
        }

        public bool Stop()
        {
            if (this.SvcManager == IntPtr.Zero)
            {
                CallSvcManager();
            }
            if (this.Service == IntPtr.Zero)
            {
                GetServiceHandle();
            }

            SERVICE_STATUS status = new SERVICE_STATUS();
            bool hResult = ControlService(this.Service, SERVICE_CONTROL.STOP, ref status);
            this.ServiceStatus = status;
            this.Wait();
            return IsStopped();
        }

        public bool Delete()
        {
            if (this.SvcManager == IntPtr.Zero)
            {
                CallSvcManager();
            }
            if (this.Service == IntPtr.Zero)
            {
                GetServiceHandle();
            }
            if (IsRunning()) {
                Console.WriteLine($"[-] Service {this.ServiceName} running! Stopping...");
                this.Stop();
                // Potentially Dangerous Operation (Infinite Loop)
                int attempts = 3;
                while (!IsStopped(true) && attempts > 0)
                {
                    this.Wait();
                    attempts -= 1;
                }
            }

            return DeleteService(this.Service);
        }

        public bool Exists()
        {
            bool exists = false;
            foreach (ServiceController service in ServiceController.GetServices(this.Target.Replace("\\\\", "")))
            {
                if (service.ServiceName == this.ServiceName)
                {
                    exists = true;
                }
            }
            return exists;
        }

        public bool IsRunning()
        {
            bool running = false;
            foreach (ServiceController service in ServiceController.GetServices(this.Target.Replace("\\\\", "")))
            {
                if ((service.ServiceName == this.ServiceName) && ((service.Status == ServiceControllerStatus.Running) || (service.Status == ServiceControllerStatus.StartPending)))
                {
                    running = true;
                }
            }
            return running;

        }

        public bool IsStopped(bool fullstop=false)
        {
            bool stopped = false;
            foreach (ServiceController service in ServiceController.GetServices(this.Target.Replace("\\\\", "")))
            {
                if ((service.ServiceName == this.ServiceName) && ((service.Status == ServiceControllerStatus.Stopped) || (service.Status == ServiceControllerStatus.StopPending)))
                {
                    stopped = (fullstop ?  true : (service.Status == ServiceControllerStatus.Stopped));
                }
            }
            return stopped;

        }

        public void GetServiceHandle()
        {
            if (this.SvcManager == IntPtr.Zero)
            {
                CallSvcManager();
            }
            if (this.Exists())
            {
                try
                {
                    this.Service = OpenService(this.SvcManager, this.ServiceName, SERVICE_ACCESS.SERVICE_ALL_ACCESS);
                }
                catch
                {
                    throw new Exception("Cannot open service");
                }
            }
            else {
                throw new Exception($"Error:  Unknown service {this.ServiceName}");
            }
        }

        public void CloseHandles()
        {
            if (this.SvcManager != null)
            {
                CloseServiceHandle(this.SvcManager);
            }
            if (this.Service != null)
            {
                CloseServiceHandle(this.Service);
            }
        }

        public void Wait()
        {
            Thread.Sleep(3000);
        }

        public ServiceExecutor(string target, string serviceName, string binPath)
        {

            this.Target = target;
            this.ServiceName = this.ServiceDisplayName = serviceName;
            this.ServiceExecutable = binPath;
            this.SvcManager = IntPtr.Zero;
            this.Service = IntPtr.Zero;
            CallSvcManager();
            //GetServiceHandle();

        }
    }
    class Program
    {
        
        static void Main(string[] args)
        {

            string FAILED_ACCESS_MANAGER = "[x] Failed to access service manager";
            string FAILED_OPEN_SERVICE = "[x] Failed to open service";
            string FAILED_CREATE_SERVICE = "[x] Failed to create service";
            string FAILED_START_SERVICE = "[x] Failed to start service";
            string FAILED_STOP_SERVICE = "[x] Failed to stop service";
            string FAILED_DELETE_SERVICE = "[x] Failed to delete service";
            string SUCCESS_ACCESS_MANAGER = "[+] Accessed service manager";
            string SUCCESS_OPEN_SERVICE = "[+] Service handle created";
            string SUCCESS_CREATE_SERVICE = "[+] Service created";
            string SUCCESS_START_SERVICE = "[+] Service started";
            string SUCCESS_STOP_SERVICE = "[+] Service stopped";
            string SUCCESS_DELETE_SERVICE = "[+] Service deleted";

            if (args.Length < 3)
            {
                Console.WriteLine(" [x] Invalid number of arguments");
                Console.WriteLine("     Usage: CsExec.exe <targetMachine> <serviceName> <binPath> <action>");
                Console.WriteLine("     Actions: start|stop|create|delete");
                Console.WriteLine("     Actions: leave blank to perform a full sequence create|start|stop|delete");
                return;
            }

            string target = $@"\\{args[0]}";
            string serviceName = args[1];
            string binPath = args[2];
            string action = args.Length < 4 ? "sequence" : args[3].ToLowerInvariant();
 
            try
            {

                ServiceExecutor sex = new ServiceExecutor(target, serviceName, binPath);
                Console.WriteLine($"[+] Created Service Executor on {target}");
                Console.Write($"Service Info:" +
                    $"{Environment.NewLine}\tService Name: {sex.ServiceName}" +
                    $"{Environment.NewLine}\tService Display Name: {sex.ServiceDisplayName}" +
                    $"{Environment.NewLine}\tBinPath: {sex.ServiceExecutable}" +
                    $"{Environment.NewLine}");

                if ((sex.SvcManager == null) || (sex.SvcManager == IntPtr.Zero)){
                    Console.WriteLine(FAILED_ACCESS_MANAGER);
                    sex.CloseHandles();
                    return;
                }
                Console.WriteLine(SUCCESS_ACCESS_MANAGER);
                if (sex.Service == null) {
                    Console.WriteLine(FAILED_OPEN_SERVICE);
                    sex.CloseHandles();
                    return;
                }
                Console.WriteLine(SUCCESS_OPEN_SERVICE);

                if (action == "sequence") 
                {  
                    if (!sex.Create()) {
                        Console.WriteLine(FAILED_CREATE_SERVICE);
                        sex.CloseHandles();
                        return;
                    }
                    Console.WriteLine(SUCCESS_CREATE_SERVICE);
                    if (!sex.Start())
                    {
                        Console.WriteLine(FAILED_START_SERVICE);
                        sex.CloseHandles();
                        return;
                    }
                    Console.WriteLine(SUCCESS_START_SERVICE);
                    if (!sex.Stop())
                    {
                        Console.WriteLine(FAILED_STOP_SERVICE);
                        sex.CloseHandles();
                        return;
                    }
                    Console.WriteLine(SUCCESS_STOP_SERVICE);
                    if (!sex.Delete())
                    {
                        Console.WriteLine(FAILED_DELETE_SERVICE);
                        sex.CloseHandles();
                        return;
                    }
                    Console.WriteLine(SUCCESS_DELETE_SERVICE);

                }
                else if (action == "create")
                {
                    if (!sex.Create())
                    {
                        Console.WriteLine(FAILED_CREATE_SERVICE);
                    }
                    Console.WriteLine(SUCCESS_CREATE_SERVICE);
                } else if (action == "start")
                {
                    if (!sex.Start())
                    {
                        Console.WriteLine(FAILED_START_SERVICE);
                    }
                    Console.WriteLine(SUCCESS_START_SERVICE);

                } else if (action == "stop")
                {
                    if (!sex.Stop())
                    {
                        Console.WriteLine(FAILED_STOP_SERVICE);
                    }
                    Console.WriteLine(SUCCESS_STOP_SERVICE);
                } else if (action == "delete")
                {
                    if (!sex.Delete())
                    {
                        Console.WriteLine(FAILED_DELETE_SERVICE);
                    }
                    Console.WriteLine(SUCCESS_DELETE_SERVICE);
                } else {
                    Console.WriteLine("[x] Invalid Action");                
                }
                sex.CloseHandles();
            }
            catch (Exception e)
            {
                Console.WriteLine("[x] {0}", e.Message);
            }
            
        }
    }
}