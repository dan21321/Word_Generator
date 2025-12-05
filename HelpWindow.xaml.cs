using System;
using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Documents;
using Word_Generator.Resources;

namespace Word_Generator
{
    /// <summary>
    /// Логика взаимодействия для HelpWindow.xaml
    /// </summary>
    public partial class HelpWindow : Window
    {
        public HelpWindow()
        {            
            InitializeComponent();
        }
        private void Email_Click(object sender, RoutedEventArgs e)
        {
            string email = "dan.sinckovckin2014@yandex.ru";

            Clipboard.SetText(email);

            MessageBox.Show(
                Strings.TxtEmailCopy,
                Strings.OK,
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );
        }
    }
}
