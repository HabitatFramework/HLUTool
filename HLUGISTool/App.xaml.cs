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

#define THREADED

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using HLU.UI.View;
using HLU.UI.ViewModel;
using HLU.Properties;

namespace HLU
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private static Thread _threadSplashScreen;
        private static ViewModelWindowSplash _splashViewModel;
        private WindowMain _mainWindow;
        private ViewModelWindowMain _mainViewModel;

        public static ViewModelWindowSplash SplashViewModel
        {
            get { return _splashViewModel; }
        }

        public static string[] StartupArguments = null;

        private void Application_Startup(object sender, StartupEventArgs e)
        {
        }

        private void Application_Exit(object sender, ExitEventArgs e)
        {
            if (_mainViewModel != null)
                _mainViewModel.OnRequestClose();
        }

        private void Application_Activated(object sender, System.EventArgs e)
        {
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            StartupArguments = e.Args;

            try
            {
#if !THREADED
                WindowSplash splashWindow = new WindowSplash();
                _splashViewModel = new ViewModelWindowSplash();
                splashWindow.DataContext = _splashViewModel;
                _splashViewModel.ProgressText = "Initialising...";
                splashWindow.Show();

                _mainWindow = new WindowMain();
                _mainViewModel = new ViewModelWindowMain();
#else

                UpgradeSettings();

                _threadSplashScreen = new Thread(ExecuteSplashScreen);
                _threadSplashScreen.SetApartmentState(ApartmentState.STA);
                _threadSplashScreen.Start();

                _mainWindow = new WindowMain();
                _mainViewModel = new ViewModelWindowMain();

                bool mainWinShowInTaskbar = _mainWindow.ShowInTaskbar;
                SizeToContent mainWinSizeToContent = _mainWindow.SizeToContent;
                _mainWindow.ShowInTaskbar = false;
                _mainWindow.SizeToContent = SizeToContent.Manual;

                _mainWindow.WindowState = WindowState.Minimized;
                _mainWindow.Show();
#endif

                DispatcherHelper.DoEvents();

                // Initialise the main view (start the tool)
                if (!_mainViewModel.Initialize())
                {
                    _mainWindow.Close();
                }
                else
                {
                    EventHandler handler = null;
                    handler = delegate
                    {
                        _mainViewModel.RequestClose -= handler;
                        _mainWindow.Close();
                    };
                    _mainViewModel.RequestClose += handler;

                    _mainWindow.DataContext = _mainViewModel;
                    App.Current.MainWindow = _mainWindow;

#if !THREADED
                    _mainWindow.Cursor = Cursors.Arrow;
                    splashWindow.Close();
                    _mainWindow.Show();
#else
                    _mainWindow.ShowInTaskbar = mainWinShowInTaskbar;
                    _mainWindow.WindowState = WindowState.Normal;
                    _mainWindow.SizeToContent = mainWinSizeToContent;
#endif

                    _mainWindow.Activate();
                }
            }
            finally
            {
                if ((_threadSplashScreen != null) && _threadSplashScreen.IsAlive)
                    _threadSplashScreen.Abort();
            }
        }

        private void UpgradeSettings()
        {
            if (Settings.Default.CallUpgrade)
            {
                Settings.Default.Upgrade();
                Settings.Default.CallUpgrade = false;
                Settings.Default.Save();
            }
        }

        private void ExecuteSplashScreen()
        {
            WindowSplash splashWindow = new WindowSplash();
            _splashViewModel = new ViewModelWindowSplash();
            splashWindow.DataContext = _splashViewModel;
            _splashViewModel.ProgressText = "Initialising...";
            splashWindow.ShowDialog();
        }

        public static Window GetActiveWindow()
        {
            if (App.Current.Windows != null)
            {
                IEnumerable<Window> appWins = App.Current.Windows.Cast<Window>();
                var q = appWins.Where(w => w.IsActive);
                if (q.Count() > 0)
                {
                    return q.ElementAt(0);
                }
                else
                {
                    q = appWins.Where(w => w.IsLoaded);
                    if (q.Count() > 0) return q.ElementAt(0);
                }
            }
            return null;
        }
    }
}
