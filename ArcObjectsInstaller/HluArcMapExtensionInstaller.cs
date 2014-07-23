// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// Copyright © 2013 Thames Valley Environmental Records Centre
// 
// This file is part of HLUTool.
// 
// HLUTool is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// HLUTool is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with HLUTool.  If not, see <http://www.gnu.org/licenses/>.

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
using ESRI.ArcGIS.ADF.CATIDs;

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
            string tlbPath = Path.Combine(Path.GetDirectoryName(regAssembly.Location),
                Path.GetFileNameWithoutExtension(regAssembly.Location) + ".tlb");

            if (File.Exists(tlbPath))
            {
                RegistrationServices regSrv = new RegistrationServices();
                regSrv.RegisterAssembly(regAssembly, AssemblyRegistrationFlags.SetCodeBase);
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

                string regKey = string.Format("HKEY_CLASSES_ROOT\\CLSID\\{{{0}}}", "c61db89f-7118-4a10-a5c1-d4a375867a02");
                if (install)
                {
                    MxExtension.Register(regKey);
                }
                else
                    MxExtension.Unregister(regKey);

                if (arcVersion > 9)
                {
                    ProcessStartInfo psi = new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(
                        Environment.SpecialFolder.CommonProgramFiles), @"ArcGIS\bin\ESRIRegasm.exe"));
                    //psi.Arguments = String.Format(@"{0} /p:Desktop{1} /s",
                    //    Path.GetFileName(base.GetType().Assembly.Location), install ? String.Empty : @" /u");

                    psi.Arguments = String.Format(@"{0} /p:Desktop{1} /s",
                         "\"" + Path.Combine(Path.GetDirectoryName(base.GetType().Assembly.Location), "HluArcMapExtension.dll") + "\"", install ? String.Empty : @" /u");

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
