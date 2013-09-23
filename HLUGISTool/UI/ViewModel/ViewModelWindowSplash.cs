using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HLU.UI.ViewModel
{
    public class ViewModelWindowSplash : ViewModelBase
    {
        private string _displayName = "HLU GIS Tool";
        private string _progressText;

        public override string DisplayName
        {
            get { return _displayName; }
            set { _displayName = value; }
        }

        public override string WindowTitle
        {
            get { return DisplayName; }
        }

        public string ProgressText
        {
            get { return _progressText; }
            set
            {
                _progressText = value;
                OnPropertyChanged("ProgressText");
            }
        }
    }
}
