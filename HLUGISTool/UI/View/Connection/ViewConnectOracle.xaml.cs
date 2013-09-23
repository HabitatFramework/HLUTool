using System;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Controls;
using System.Windows.Data;

namespace HLU.UI.View.Connection
{
    /// <summary>
    /// Interaction logic for ViewConnectOracle.xaml
    /// </summary>
    public partial class ViewConnectOracle : Window
    {
        IntPtr _windowHandle;

        public ViewConnectOracle()
        {
            InitializeComponent();
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            HwndSource hwndSrc = PresentationSource.FromVisual(this) as HwndSource;
            if (hwndSrc != null) _windowHandle = hwndSrc.Handle;

            if (this.ComboBoxDataSource.Items.Count == 1) this.ComboBoxDataSource.SelectedIndex = 0;
        }

        private void ComboBox_DropDownOpened(object sender, EventArgs e)
        {
            Binding bnd = BindingOperations.GetBinding((ComboBox)sender, ComboBox.SelectedItemProperty);
            if (bnd != null)
                ((HLU.UI.ViewModel.ViewModelConnectOracle)this.DataContext).ViewEvents(_windowHandle, bnd.Path.Path);
            else
                ((HLU.UI.ViewModel.ViewModelConnectOracle)this.DataContext).ViewEvents(_windowHandle, null);
        }
    }
}
