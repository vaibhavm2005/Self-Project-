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
using static demo.Window1;

namespace demo
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
     
            public Window1()
            {
                InitializeComponent();
            }

        private void btnLogin_Click(object sender, RoutedEventArgs e)
        {
                // Hardcoded username and password for demonstration
                string username = "admin";
                string password = "1234";

                if (txtUsername.Text == username && txtPassword.Text == password)
                {
                    MessageBox.Show("Login successful!", "Success");
                    // Open Form2 (Main Form)
                    Window2 mainForm = new Window2();
                    mainForm.Show();
                    this.Hide(); // Hide the login form
                }
                else
                {
                    MessageBox.Show("Invalid username or password", "Error");
                }
            }

        
    }
}
  