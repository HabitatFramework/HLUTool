using System.Windows;

namespace HLU.UI.View
{
    /// <summary>
    /// Interaction logic for MessageWindow.xaml
    /// </summary>
    public partial class MessageWindow : Window
    {
        public MessageWindow()
        {
            InitializeComponent();
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (e.HeightChanged && e.PreviousSize.Height != 0)
                this.TextBoxMessage.Height = e.NewSize.Height - (e.PreviousSize.Height - this.TextBoxMessage.Height);
        }
    }
}
