using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;

namespace HLU.UI.View.Connection
{
    /// <summary>
    /// Interaction logic for ViewConnectOdbc.xaml
    /// </summary>
    public partial class ViewConnectOdbc : Window
    {
        private IntPtr _windowHandle;

        public ViewConnectOdbc()
        {
            InitializeComponent();
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            HwndSource hwndSrc = PresentationSource.FromVisual(this) as HwndSource;
            if (hwndSrc != null) _windowHandle = hwndSrc.Handle;

            ((HLU.UI.ViewModel.ViewModelConnectOdbc)this.DataContext).ViewEvents(_windowHandle, null);

            if ((this.ComboBoxDsn.Items.Count == 1) && (String.IsNullOrEmpty(this.ComboBoxDsn.Text) ||
                this.ComboBoxDsn.Items.Contains(this.ComboBoxDsn.Text))) this.ComboBoxDsn.SelectedIndex = 0;
        }

        private void ComboBox_DropDownOpened(object sender, EventArgs e)
        {
            Binding bnd = BindingOperations.GetBinding((ComboBox)sender, ComboBox.SelectedItemProperty);
            if (bnd != null)
                ((HLU.UI.ViewModel.ViewModelConnectOdbc)this.DataContext).ViewEvents(_windowHandle, bnd.Path.Path);
            else
                ((HLU.UI.ViewModel.ViewModelConnectOdbc)this.DataContext).ViewEvents(_windowHandle, null);
        }
    }
}
