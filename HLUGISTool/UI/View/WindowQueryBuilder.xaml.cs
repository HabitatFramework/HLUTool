using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace HLU.UI.View
{
    /// <summary>
    /// Interaction logic for WindowQueryBuilder.xaml
    /// </summary>
    public partial class WindowQueryBuilder : Window
    {
        public WindowQueryBuilder()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Sets the text property of a DataGridComboBoxColumnWithBinding when user has re-entered it and 
        /// edited its text portion.
        /// Used by the query builder's Value column.
        /// </summary>
        /// <param name="sender">ComboBox control.</param>
        /// <param name="e"></param>
        private void ComboBox_PreviewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            SetComboBoxText(sender as ComboBox);
        }

        /// <summary>
        /// Sets the text property of a DataGridComboBoxColumnWithBinding when user has re-entered it and 
        /// edited its text portion.
        /// Used by the query builder's Value column.
        /// </summary>
        /// <param name="cmb">ComboBox control.</param>
        private void SetComboBoxText(ComboBox cmb)
        {
            if (cmb != null)
            {
                string text = cmb.Text;
                Binding bnd = BindingOperations.GetBinding(cmb, ComboBox.TextProperty);
                if (bnd != null)
                {
                    PropertyInfo pi = cmb.DataContext.GetType().GetProperty(bnd.Path.Path);
                    if (pi != null)
                    {
                        object propValue = cmb.DataContext.GetType().GetProperty(bnd.Path.Path).GetValue(cmb.DataContext, null);
                        if ((propValue != null) && !propValue.Equals(text))
                            pi.SetValue(cmb.DataContext, text, null);
                    }
                }
            }
        }
    }
}
