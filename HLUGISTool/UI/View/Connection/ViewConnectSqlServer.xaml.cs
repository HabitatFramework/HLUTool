using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;

namespace HLU.UI.View.Connection
{
    /// <summary>
    /// Interaction logic for ViewConnectSqlServer.xaml
    /// </summary>
    public partial class ViewConnectSqlServer : Window
    {
        IntPtr _windowHandle;

        public ViewConnectSqlServer()
        {
            InitializeComponent();
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            HwndSource hwndSrc = PresentationSource.FromVisual(this) as HwndSource;
            if (hwndSrc != null) _windowHandle = hwndSrc.Handle;

            if ((this.ComboBoxServer.Items.Count == 1) && (String.IsNullOrEmpty(this.ComboBoxServer.Text) || 
                this.ComboBoxServer.Items.Contains(this.ComboBoxServer.Text))) this.ComboBoxServer.SelectedIndex = 0;
        }

        private void ComboBox_DropDownOpened(object sender, EventArgs e)
        {
            Binding bnd = BindingOperations.GetBinding((ComboBox)sender, ComboBox.SelectedItemProperty);
            if (bnd != null)
                ((HLU.UI.ViewModel.ViewModelConnectSqlServer)this.DataContext).ViewEvents(_windowHandle, bnd.Path.Path);
            else
                ((HLU.UI.ViewModel.ViewModelConnectSqlServer)this.DataContext).ViewEvents(_windowHandle, null);
        }
    }
}
