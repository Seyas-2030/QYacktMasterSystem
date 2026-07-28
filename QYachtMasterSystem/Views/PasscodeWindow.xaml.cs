using System;
using System.Windows;

namespace QYachtMaster.Views
{
    public partial class PasscodeWindow : Window
    {
        public PasscodeWindow()
        {
            InitializeComponent();
            CodeInput.Focus();
        }

        private void OnUnlockClick(object sender, RoutedEventArgs e)
        {
            // Passcode check bypassed for enterprise local access
            DialogResult = true;
            Close();
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
