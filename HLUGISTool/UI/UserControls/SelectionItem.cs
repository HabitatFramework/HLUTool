using System;
using System.ComponentModel;

namespace HLU.UI.UserControls
{
    public class SelectionItem<T> : INotifyPropertyChanged
    {
        #region Fields

        private bool isSelected;

        private T item;

        #endregion

        #region Properties

        public bool IsSelected
        {
            get { return isSelected; }
            set
            {
                if (value == isSelected) return;
                isSelected = value;
                OnPropertyChanged("IsSelected");
                OnSelectionChanged();
            }
        }

        public T Item
        {
            get { return item; }
            set
            {
                if (value.Equals(item)) return;
                item = value;
                OnPropertyChanged("Item");
            }
        }

        #endregion

        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        public event EventHandler SelectionChanged;

        #endregion

        #region Constructor

        public SelectionItem(T item)
            : this(false, item)
        {
        }

        public SelectionItem(bool selected, T item)
        {
            this.isSelected = selected;
            this.item = item;
        }

        #endregion

        #region Event invokers

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler changed = PropertyChanged;
            if (changed != null) changed(this, new PropertyChangedEventArgs(propertyName));
        }

        private void OnSelectionChanged()
        {
            EventHandler changed = SelectionChanged;
            if (changed != null) changed(this, EventArgs.Empty);
        }

        #endregion
    }
}
