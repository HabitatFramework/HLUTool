using System;
using System.IO;
using System.Windows;
using HLU.GISApplication.ArcGIS;
using HLU.GISApplication.MapInfo;
using HLU.Properties;
using HLU.UI.View;
using HLU.UI.ViewModel;
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
        private static Nullable<bool> _arcGisInstalled;
        private static Nullable<bool> _mapInfoInstalled;
        private static WindowSelectGIS _windowSelGIS;
        private static ViewModelWindowSelectGIS _viewModelSelGIS;
        private static GISApplications _gisApp;
        private static bool _cancelled;

        public static GISApp CreateGisApp()
        {
            try
            {
                _gisApp = GISApplications.None;

                if (Enum.IsDefined(typeof(GISApplications), Settings.Default.PreferredGis))
                    _gisApp = (GISApplications)Settings.Default.PreferredGis;

                if (_gisApp == GISApplications.None)
                {
                    if (ArcGisInstalled && MapInfoInstalled)
                    {
                        SelectGISApp();
                    }
                    else if (ArcGisInstalled)
                    {
                        _gisApp = GISApplications.ArcGIS;
                    }
                    else if (MapInfoInstalled)
                    {
                        _gisApp = GISApplications.MapInfo;
                    }

                    Settings.Default.PreferredGis = (int)_gisApp;
                }

                if (_gisApp == GISApplications.None)
                    throw new ArgumentException("Could not find GIS application.");
                else
                    Settings.Default.Save();

                switch (_gisApp)
                {
                    case GISApplications.ArcGIS:
                        return new ArcMapApp(Settings.Default.MapPath);
                    case GISApplications.MapInfo:
                        return new MapInfoApp(Settings.Default.MapPath);
                    default:
                        return null;
                }
            }
            catch (Exception ex)
            {
                if (!_cancelled)
                    MessageBox.Show(ex.Message, "HBIC: Fatal Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        public static bool ArcGisInstalled
        {
            get
            {
                if (_arcGisInstalled == null)
                    _arcGisInstalled = Type.GetTypeFromProgID("esriEditor.editor", false) != null;
                return (bool)_arcGisInstalled;
            }
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

        private static void SelectGISApp()
        {
            _windowSelGIS = new WindowSelectGIS();
            _windowSelGIS.Owner = App.Current.MainWindow;
            _windowSelGIS.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            _viewModelSelGIS = new ViewModelWindowSelectGIS();
            _viewModelSelGIS.RequestClose +=
                new ViewModelWindowSelectGIS.RequestCloseEventHandler(_viewModelSelGIS_RequestClose);

            _windowSelGIS.DataContext = _viewModelSelGIS;

            _windowSelGIS.ShowDialog();
        }

        private static void _viewModelSelGIS_RequestClose(bool cancelled, GISApplications selectedGISApp)
        {
            _viewModelSelGIS.RequestClose -= _viewModelSelGIS_RequestClose;
            _windowSelGIS.Close();

            _cancelled = cancelled;
            if (!_cancelled) _gisApp = selectedGISApp;
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
