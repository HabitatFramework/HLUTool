using System;
using System.Windows;
using HLU.UI.View;
using HLU.UI.ViewModel;

namespace HLU
{
    internal static class ShowMessageWindow
    {
        private static MessageWindow _messageWindow;
        private static ViewModelMessageWindow _messageWindowViewModel;

        internal static void ShowMessage(string messageText, string messageHeader)
        {
            _messageWindow = new MessageWindow();
            if ((_messageWindow.Owner = App.GetActiveWindow()) == null)
                throw (new Exception("No parent window loaded"));

            _messageWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;

            // Create ViewModel to which main window binds
            _messageWindowViewModel = new ViewModelMessageWindow();
            _messageWindowViewModel.MessageText = messageText;
            _messageWindowViewModel.MessageHeader = messageHeader;

            // When ViewModel asks to be closed, close window
            _messageWindowViewModel.RequestClose += new EventHandler(CloseMessageWindow);

            // Allow all controls in window to bind to ViewModel by setting DataContext
            _messageWindow.DataContext = _messageWindowViewModel;

            // Show window
            _messageWindow.ShowDialog();
        }

        /// <summary>
        /// Closes help window and removes close window handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        internal static void CloseMessageWindow(object sender, EventArgs e)
        {
            _messageWindowViewModel.RequestClose -= CloseMessageWindow;
            _messageWindow.Close();
        }
    }
}
