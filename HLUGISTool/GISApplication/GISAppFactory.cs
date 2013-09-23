using System;
using System.IO;
using System.Windows;
using HLU.GISApplication.MapInfo;
using HLU.Properties;
using Microsoft.Win32;

namespace HLU.GISApplication
{
    public enum GISApplications
    {
        None,
        ArcGIS,
        MapInfo
    };

    class GISAppFactory
    {
        private static Nullable<bool> _mapInfoInstalled;

        public static GISApp CreateGisApp()
        {
            try
            {
                GISApplications gisApp = GISApplications.None;

                if (Enum.IsDefined(typeof(GISApplications), Settings.Default.PreferredGis))
                    gisApp = (GISApplications)Settings.Default.PreferredGis;

                if (gisApp == GISApplications.None)
                {
                    if (MapInfoInstalled)
                    {
                        gisApp = GISApplications.MapInfo;
                    }

                    Settings.Default.PreferredGis = (int)gisApp;
                }

                if (gisApp == GISApplications.None)
                    throw new ArgumentException("Could not find GIS application.");
                else
                    Settings.Default.Save();

                switch (gisApp)
                {
                    case GISApplications.MapInfo:
                        return new MapInfoApp(Settings.Default.MapPath);
                    default:
                        return null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "HBIC: Fatal Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public static bool ArcGisInstalled
        {
            get { return false; }
        }

        public static bool MapInfoInstalled
        {
            get
            {
                if (_mapInfoInstalled == null)
                    _mapInfoInstalled = Type.GetTypeFromProgID("MapInfo.Application", false) != null;
                return (bool)_mapInfoInstalled;
            }
        }

        public static string GetMapPath(GISApplications gisApp)
        {
            try
            {
                string mapPath = Settings.Default.MapPath;

                if (File.Exists(mapPath))
                {
                    FileInfo mapFile = new FileInfo(mapPath);
                    switch (gisApp)
                    {
                        case GISApplications.ArcGIS:
                            if (mapFile.Extension.ToLower() == ".mxd") return mapPath;
                            break;
                        case GISApplications.MapInfo:
                            if (mapFile.Extension.ToLower() == ".wor") return mapPath;
                            break;
                    }
                }

                OpenFileDialog openFileDlg = new OpenFileDialog();
                switch (gisApp)
                {
                    case GISApplications.ArcGIS:
                        openFileDlg.Filter = "ESRI ArcMap Documents (*.mxd)|*.mxd";
                        openFileDlg.Title = "Open HLU Map Document";
                        break;
                    case GISApplications.MapInfo:
                        openFileDlg.Filter = "MapInfo Workspaces (*.wor)|*.wor";
                        openFileDlg.Title = "Open HLU Workspace";
                        break;
                    default:
                        return null;
                }
                openFileDlg.Multiselect = false;
                openFileDlg.CheckPathExists = true;
                openFileDlg.CheckFileExists = true;
                openFileDlg.ValidateNames = true;
                openFileDlg.RestoreDirectory = false;
                openFileDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (openFileDlg.ShowDialog() == true)
                {
                    if (File.Exists(openFileDlg.FileName))
                        Settings.Default.MapPath = openFileDlg.FileName;
                    return openFileDlg.FileName;
                }
            }
            catch { }

            return null;
        }

        public static bool ClearSettings()
        {
            try
            {
                Settings.Default.PreferredGis = (int)GISApplications.None;
                Settings.Default.MapPath = String.Empty;
                Settings.Default.Save();

                return true;
            }
            catch { return false; }
        }
    }
}
