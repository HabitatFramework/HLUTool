using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;

namespace HLU.UI.View.Connection
{
    /// <summary>
    /// Interaction logic for ViewConnectOleDb.xaml
    /// </summary>
    public partial class ViewConnectOleDb : Window
    {
        private IntPtr _windowHandle;

        public ViewConnectOleDb()
        {
            InitializeComponent();
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            HwndSource hwndSrc = PresentationSource.FromVisual(this) as HwndSource;
            if (hwndSrc != null) _windowHandle = hwndSrc.Handle;

            ((HLU.UI.ViewModel.ViewModelConnectOleDb)this.DataContext).ViewEvents(_windowHandle, null);
        }

        private void ComboBox_DropDownOpened(object sender, EventArgs e)
        {
            Binding bnd = BindingOperations.GetBinding((ComboBox)sender, ComboBox.SelectedItemProperty);
            if (bnd != null)
                ((HLU.UI.ViewModel.ViewModelConnectOleDb)this.DataContext).ViewEvents(_windowHandle, bnd.Path.Path);
            else
                ((HLU.UI.ViewModel.ViewModelConnectOleDb)this.DataContext).ViewEvents(_windowHandle, null);
        }
    }
}
