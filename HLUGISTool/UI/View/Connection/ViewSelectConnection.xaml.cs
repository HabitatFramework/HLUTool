using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace HLU.UI.View.Connection
{
    /// <summary>
    /// Interaction logic for ViewSelectConnection.xaml
    /// </summary>
    public partial class ViewSelectConnection : Window
    {
        IntPtr _windowHandle;

        public ViewSelectConnection()
        {
            InitializeComponent();
            Loaded += (sender, e) => MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            HwndSource hwndSrc = PresentationSource.FromVisual(this) as HwndSource;
            if (hwndSrc != null) _windowHandle = hwndSrc.Handle;
        }
    }
}
