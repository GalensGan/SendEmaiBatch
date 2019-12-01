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

namespace SendMailBatch
{
    /// <summary>
    /// LoginWpf.xaml 的交互逻辑
    /// </summary>
    public partial class LoginWpf : Window
    {
        public LoginWpf()
        {
            InitializeComponent();
            passwordBox.KeyDown += PasswordBox_KeyDown;
        }

        private void PasswordBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.Hide();
                string key = Encryption.Get32MD5One(passwordBox.Password);
                MainWindow f = new MainWindow(key);
                f.Show();
                this.Owner = f;
                this.Hide();
            }
        }
    }
}
