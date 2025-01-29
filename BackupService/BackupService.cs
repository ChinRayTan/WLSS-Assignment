using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Reflection;
using System.ServiceProcess;
using System.Windows;
using System.Xml.Linq;
using static BackupService.PInvokes;

namespace BackupService
{
    public partial class BackupService : ServiceBase
    {
        private System.Timers.Timer timer = new System.Timers.Timer();
        private static XDocument settings = XDocument.Load(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Settings.xml");

        private class Settings
        {
            public readonly static DateTime BackupTime = DateTime.Today.AddDays(1).AddHours(Convert.ToDouble(settings.Root.Element("BackupTime").Value));
            public readonly static string TargetShare = settings.Root.Element("TargetShare").Value.ToString();
            public readonly static string Username = settings.Root.Element("Username").Value.ToString();
            public readonly static string Password = settings.Root.Element("Password").Value.ToString();
        }

        public BackupService()
        {
            InitializeComponent();
            eventLog1 = new EventLog();
            if (!EventLog.SourceExists("WLSS Backup")) EventLog.CreateEventSource("WLSS Backup", "Logs");
            eventLog1.Source = "WLSS Backup";
            eventLog1.Log = "Logs";
        }

        protected override void OnStart(string[] args)
        {
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            eventLog1.WriteEntry($"Service started at {DateTime.Now}.", EventLogEntryType.Information);
            timer.Interval = Settings.BackupTime.Subtract(DateTime.Now).TotalSeconds * 1000;
            timer.Elapsed += Timer_Elapsed;
            timer.Start();

            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            ExecuteBackup();
        }

        protected override void OnStop()
        {
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOP_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            eventLog1.WriteEntry($"Service stopped at {DateTime.Now}.", EventLogEntryType.Information);

            serviceStatus.dwCurrentState = ServiceState.SERVICE_STOPPED;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            ExecuteBackup();
        }

        private void ExecuteBackup()
        {
            try
            {
                timer.Interval = 24 * 60 * 60 * 1000;

                eventLog1.WriteEntry("Time to backup - Beginning backup now.");
                eventLog1.WriteEntry("Now processing: Connecting to backup share.");

                Process mountNetworkDrive = new Process();
                mountNetworkDrive.StartInfo.FileName = "net.exe";
                mountNetworkDrive.StartInfo.Arguments = $@"use Z: \\{Settings.TargetShare} /user:{Settings.Username} {Settings.Password}";
                mountNetworkDrive.StartInfo.UseShellExecute = false;
                mountNetworkDrive.StartInfo.RedirectStandardOutput = true;
                mountNetworkDrive.StartInfo.CreateNoWindow = true;
                mountNetworkDrive.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                mountNetworkDrive.Start();
                mountNetworkDrive.WaitForExit();

                eventLog1.WriteEntry("Now processing: Creating text file on backup share.");

                RegistryKey hklmApps = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", true);
                using (StreamWriter sw = new StreamWriter($@"Z:\AppsList_{DateTime.Today.ToString("yyyy-MM-dd")}"))
                {
                    string header = $"List of apps on {Environment.MachineName} as of {DateTime.Today.ToString("g")}";
                    sw.WriteLine(header);
                    sw.WriteLine(new String('-', header.Length));
                    sw.WriteLine();
                    sw.WriteLine();
                    sw.WriteLine("Apps in HKLM:");
                    sw.WriteLine();

                    eventLog1.WriteEntry("Now processing: Apps in HKLM.");
                    foreach (string registryKey in hklmApps.GetSubKeyNames())
                    {
                        WriteDataToFile(sw, hklmApps, registryKey);
                    }

                    RegistryKey users = Registry.Users;
                    foreach (string user in users.GetSubKeyNames())
                    {
                        if (users.OpenSubKey(user).OpenSubKey(@"SOFTWARE\Microsoft\Windows\Uninstall") == null) continue;
                        else
                        {
                            eventLog1.WriteEntry($"Now processing: Apps in user {Registry.LocalMachine.OpenSubKey($@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\{user}").GetValue("ProfileImagePath").ToString().Split('\\')[2]}");
                            RegistryKey userApps = users.OpenSubKey(user).OpenSubKey(@"SOFTWARE\Microsoft\Windows\Uninstall");
                            foreach (string registryKey in userApps.GetSubKeyNames())
                            {
                                WriteDataToFile(sw, userApps, registryKey);
                            }
                            sw.WriteLine();
                        }
                    }

                }

                eventLog1.WriteEntry("Backup finished. See you tomorrow, same time.");
            } catch (Exception ex)
            {
                OnException(ex);
            }
        }

        private void WriteDataToFile(StreamWriter sw, RegistryKey parentRegistryKey, string childRegistryKey)
        {
            RegistryKey subKey = parentRegistryKey.OpenSubKey(childRegistryKey);
            sw.WriteLine($"Name: {subKey.GetValue("DisplayName")}");
            sw.WriteLine($"Publisher: {subKey.GetValue("Publisher")}");
            sw.WriteLine($"Version: {subKey.GetValue("DisplayVersion")}");
            sw.WriteLine($"Install date: {Convert.ToDateTime(subKey.GetValue("InstallDate")).ToString("d")}");

            if (string.IsNullOrEmpty(subKey.GetValue("InstallLocation").ToString()))
            {
                ManagementObjectSearcher mos = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Product");
                foreach (ManagementObject mo in mos.Get())
                {
                    if (mo["Name"].ToString().Contains(subKey.GetValue("DisplayName").ToString()))
                    {
                        sw.WriteLine($"Install location: {mo["InstallLocation"]}");
                        break;
                    }
                }
            }
            else
            {
                sw.WriteLine($"Install location: {subKey.GetValue("InstallLocation")}");
            }
        }

        private void OnException(Exception ex)
        {
            eventLog1.WriteEntry(ex.Message, EventLogEntryType.Error);
        }
    }
}
