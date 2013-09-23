using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Reflection;

namespace HLU
{
    [RunInstaller(true)]
    public partial class HluArcMapExtensionInstaller : Installer
    {
        public HluArcMapExtensionInstaller()
        {
            InitializeComponent();
        }

        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);

            Assembly regAssembly = base.GetType().Assembly;
            string installDir = Path.GetDirectoryName(regAssembly.Location);

            if (File.Exists(Path.Combine(installDir, Path.GetFileNameWithoutExtension(regAssembly.Location) + ".tlb")))
            {
                RegistrationServices regSrv = new RegistrationServices();
                regSrv.RegisterAssembly(base.GetType().Assembly,
                    AssemblyRegistrationFlags.SetCodeBase);
            }
            else
            {
                Regasm(true, regAssembly.Location);
            }
            EsriRegasm(true);
        }

        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

            Assembly regAssembly = base.GetType().Assembly;
            string tlbPath = Path.Combine(Path.GetDirectoryName(regAssembly.Location),
                Path.GetFileNameWithoutExtension(regAssembly.Location) + ".tlb");

            RegistrationServices regSrv = new RegistrationServices();
            regSrv.UnregisterAssembly(regAssembly);

            EsriRegasm(false);
            if (File.Exists(tlbPath)) Regasm(false, regAssembly.Location);
        }


        private bool Regasm(bool register, string file)
        {
            if (String.IsNullOrEmpty(file)) throw new InstallException("Assembly not defined");
            if (!File.Exists(file)) throw new InstallException("Assembly not found");

            string cmd = Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "regasm.exe");
            string fileDir = Path.GetDirectoryName(file);
            string fileName = Path.GetFileName(file);
            string tlbName = Path.GetFileNameWithoutExtension(file) + ".tlb";

            ProcessStartInfo psi = new ProcessStartInfo(cmd);
            psi.WorkingDirectory = fileDir;
            psi.CreateNoWindow = true;
            psi.UseShellExecute = false;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;

            if (register)
            {
                psi.Arguments = String.Format("{0} /register /codebase /tlb:{1}", Quote(fileName), Quote(tlbName));
                Process regProc = new Process();
                regProc.StartInfo = psi;
                regProc.Start();
                regProc.WaitForExit();

            }
            else
            {
                string tlbPath = Path.Combine(fileDir, tlbName);
                try
                {
                    psi.Arguments = String.Format("{0} /unregister /tlb:{1}", Quote(fileName), Quote(tlbName));
                    Process regProc = new Process();
                    regProc.StartInfo = psi;
                    regProc.Start();
                    regProc.WaitForExit();
                }
                catch { return false; }
                finally { if (File.Exists(tlbPath)) File.Delete(tlbPath); }
            }

            return true;
        }

        private void EsriRegasm(bool install)
        {
            try
            {
                int arcVersion = -1;
                string regCmd = String.Empty, args = String.Empty;

                RegistryKey rk = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ESRI\ArcGIS");
                object rkVal = rk.GetValue("RealVersion");
                if ((rkVal == null) || !Int32.TryParse(rkVal.ToString().Split('.')[0], out arcVersion)) arcVersion = -1;

                if (arcVersion > 9)
                {
                    ProcessStartInfo psi = new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(
                        Environment.SpecialFolder.CommonProgramFiles), @"ArcGIS\bin\ESRIRegasm.exe"));
                    psi.Arguments = String.Format(@"{0} /p:Desktop{1} /s",
                        Path.GetFileName(base.GetType().Assembly.Location), install ? String.Empty : @" /u");
                    psi.WorkingDirectory = Path.GetDirectoryName(base.GetType().Assembly.Location);
                    psi.CreateNoWindow = true;
                    psi.UseShellExecute = false;

                    Process p = new Process();
                    p.StartInfo = psi;
                    p.Start();
                    p.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("{0}\n{1}", ex.Source, ex.Message), "HLU GIS Tool Installer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string Quote(string s)
        {
            return String.Format("{0}{1}{0}", "\"", s);
        }
    }
}
