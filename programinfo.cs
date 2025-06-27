using System;
using System.Windows.Forms;
using System.IO;
using System.Reflection;

namespace ice_cream
{
    public partial class programinfo : Form
    {
        public programinfo(string fullName = "")
        {
            InitializeComponent();
            Load += (sender, e) => {
                string appPath = Assembly.GetExecutingAssembly().Location;
                label19.Text = $"{Path.GetDirectoryName(appPath)}";

                label20.Text = string.IsNullOrEmpty(UserSession.FullName) ?
                    "Не авторизован" :
                    $"{UserSession.FullName}";
            };
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            system_requirements form2 = new system_requirements();
            form2.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            contact_information form2 = new contact_information();
            form2.Show();
        }
    }
}