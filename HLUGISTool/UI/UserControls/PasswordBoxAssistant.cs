using System;
using System.Windows;
using System.Windows.Controls;

namespace HLU.UI.UserControls
{
    static class PasswordBoxAssistant
    {
        #region DependencyProperties

        public static readonly DependencyProperty BoundPassword =
            DependencyProperty.RegisterAttached("BoundPassword", typeof(string), typeof(PasswordBoxAssistant), 
            new FrameworkPropertyMetadata(String.Empty, OnBoundPasswordChanged));

        public static readonly DependencyProperty BindPassword = DependencyProperty.RegisterAttached(
            "BindPassword", typeof(bool), typeof(PasswordBoxAssistant), new PropertyMetadata(false, OnBindPasswordChanged));

        private static readonly DependencyProperty UpdatingPassword =
            DependencyProperty.RegisterAttached("UpdatingPassword", typeof(bool), typeof(PasswordBoxAssistant));

        #endregion

        #region Public Methods

        public static void SetBindPassword(DependencyObject dp, bool value)
        {
            dp.SetValue(BindPassword, value);
        }

        public static bool GetBindPassword(DependencyObject dp)
        {
            return (bool)dp.GetValue(BindPassword);
        }

        public static string GetBoundPassword(DependencyObject dp)
        {
            return (string)dp.GetValue(BoundPassword);
        }

        public static void SetBoundPassword(DependencyObject dp, string value)
        {
            dp.SetValue(BoundPassword, value);
        }

        #endregion

        #region Private Methods

        private static void OnBoundPasswordChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            PasswordBox box = d as PasswordBox;

            // Only handle this event when the property is attached to a PasswordBox
            // and when the BindPassword attached property has been set to true
            if ((d == null) || !GetBindPassword(d)) return;

            // Avoid recursive updating by ignoring the box's changed event
            box.PasswordChanged -= HandlePasswordChanged;

            string newPassword = (string)e.NewValue;

            if (!GetUpdatingPassword(box)) box.Password = newPassword;

            box.PasswordChanged += HandlePasswordChanged;
        }

        /// <summary>
        /// When BindPassword attached property is set on a PasswordBox start listening to its PasswordChanged event
        /// </summary>
        /// <param name="dp"></param>
        /// <param name="e"></param>
        private static void OnBindPasswordChanged(DependencyObject dp, DependencyPropertyChangedEventArgs e)
        {
            PasswordBox box = dp as PasswordBox;
            if (box == null) return;

            bool wasBound = (bool)e.OldValue;
            bool needToBind = (bool)e.NewValue;

            if (wasBound) box.PasswordChanged -= HandlePasswordChanged;

            if (needToBind) box.PasswordChanged += HandlePasswordChanged;
        }

        private static void HandlePasswordChanged(object sender, RoutedEventArgs e)
        {
            PasswordBox box = sender as PasswordBox;

            // Set flag to indicate we're updating password
            SetUpdatingPassword(box, true);

            // Push new password into BoundPassword property
            SetBoundPassword(box, box.Password);
            SetUpdatingPassword(box, false);
        }

        private static bool GetUpdatingPassword(DependencyObject dp)
        {
            return (bool)dp.GetValue(UpdatingPassword);
        }

        private static void SetUpdatingPassword(DependencyObject dp, bool value)
        {
            dp.SetValue(UpdatingPassword, value);
        }

        #endregion
    }
}
