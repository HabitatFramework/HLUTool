using System.Windows;
using HLU.UI.UserControls;
using System.Windows.Controls;
using System.Collections.Generic;

namespace HLU.UI.View
{
    /// <summary>
    /// Interaction logic for WindowOptions.xaml
    /// </summary>
    public partial class WindowOptions : Window
    {
        public WindowOptions()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Fakes a click on each history column list item so that their checkboxes are properly initialised.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Activated(object sender, System.EventArgs e)
        {
            List<DependencyObject> listBoxes = new List<DependencyObject>();
            FindControls.GetChildren(this, typeof(ListBox), ref listBoxes);

            foreach (ListBox l in listBoxes)
            {
                if ((l.Items.Count > 0) && 
                    (l.Items[0].GetType().ToString().StartsWith(this.GetType().Namespace.Replace("View", "UserControls") + ".SelectionItem")))
                    ForceIsSelectedUpdate(l);
            }


        }

        private void ForceIsSelectedUpdate(ListBox listBoxControl)
        {
            for (int i = 0; i < listBoxControl.Items.Count; i++)
            {
                SelectionItem<string> it = listBoxControl.Items[i] as SelectionItem<string>;
                if (it.IsSelected)
                {
                    it.IsSelected = !it.IsSelected;
                    it.IsSelected = !it.IsSelected;
                    listBoxControl.SelectedItems.Add(listBoxControl.Items[i]);
                }
            }
        }
    }
}
