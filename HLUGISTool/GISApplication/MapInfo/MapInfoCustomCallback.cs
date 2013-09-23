using System;
using System.Runtime.InteropServices;
using MapinfoWrapper;

namespace HLU.GISApplication.MapInfo
{
    [ComVisible(true)]
    public class MapInfoCustomCallback : MapinfoCallback
    {
        public event Action<string> OnMenuItemClick;

        public void MenuItemHandler(string command)
        {
            // Store the event locally to save against a race condition.
            Action<string> menuEvent = OnMenuItemClick;
            if (menuEvent != null)
            {
                // Raise the event.
                menuEvent(command);
            }
        }
    }
}
