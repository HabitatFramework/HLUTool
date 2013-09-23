using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;

namespace HLU.UI.View.Connection
{
    /// <summary>
    /// Interaction logic for ViewConnectPgSql.xaml
    /// </summary>
    public partial class ViewConnectPgSql : Window
    {
        IntPtr _windowHandle;

        public ViewConnectPgSql()
        {
            InitializeComponent();
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            HwndSource hwndSrc = PresentationSource.FromVisual(this) as HwndSource;
            if (hwndSrc != null) _windowHandle = hwndSrc.Handle;

            if ((this.ComboBoxEncoding.Items.Count == 1) && (String.IsNullOrEmpty(this.ComboBoxEncoding.Text) ||
                this.ComboBoxEncoding.Items.Contains(this.ComboBoxEncoding.Text))) this.ComboBoxEncoding.SelectedIndex = 0;
        }

        private void ComboBox_DropDownOpened(object sender, EventArgs e)
        {
            Binding bnd = BindingOperations.GetBinding((ComboBox)sender, ComboBox.SelectedItemProperty);
            if (bnd != null)
                ((HLU.UI.ViewModel.ViewModelConnectPgSql)this.DataContext).ViewEvents(_windowHandle, bnd.Path.Path);
            else
                ((HLU.UI.ViewModel.ViewModelConnectPgSql)this.DataContext).ViewEvents(_windowHandle, null);
        }
    }
}
