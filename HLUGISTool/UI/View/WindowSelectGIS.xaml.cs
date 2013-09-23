using System.Windows;
using System.Windows.Input;

namespace HLU.UI.View
{
    /// <summary>
    /// Interaction logic for WindowSelectGIS.xaml
    /// </summary>
    public partial class WindowSelectGIS : Window
    {
        public WindowSelectGIS()
        {
            InitializeComponent();
            Loaded += (sender, e) => MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
        }
    }
}
