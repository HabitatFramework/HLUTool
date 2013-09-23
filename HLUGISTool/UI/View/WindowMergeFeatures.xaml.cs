using System.Windows;
using System.Windows.Input;

namespace HLU.UI.View
{
    /// <summary>
    /// Interaction logic for MergeFeaturesWindow.xaml
    /// </summary>
    public partial class WindowMergeFeatures : Window
    {
        public WindowMergeFeatures()
        {
            InitializeComponent();
        }

        private void DataGridQueryBuilder_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (this.ButtonFlashFeature.Command.CanExecute(null))
                this.ButtonFlashFeature.Command.Execute(null);
        }
    }
}
