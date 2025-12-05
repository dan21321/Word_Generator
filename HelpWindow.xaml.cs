using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace generateDocs
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
                "Почта скопирована в буфер обмена!",
                "Готово",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );
        }
    }
}
