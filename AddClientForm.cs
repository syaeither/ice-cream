using Npgsql;
using System;
using System.Windows.Forms;

namespace ice_cream
{
    public partial class AddClientForm : Form
    {
        public AddClientForm()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e) //Добавление клиента
        {
            string fullName = textBox1.Text;
            string phoneNumber = maskedTextBox1.Text;

            if (string.IsNullOrWhiteSpace(fullName) || string.IsNullOrWhiteSpace(phoneNumber))
            {
                MessageBox.Show("Пожалуйста, заполните все поля.");
                return;
            }

            using (NpgsqlConnection connection = new NpgsqlConnection(Program.ConnectionString))
            {
                try
                {
                    connection.Open();
                    using (NpgsqlCommand command = new NpgsqlCommand("INSERT INTO clients (full_name, telephone, created_at) VALUES (@fullName, @phoneNumber, CURRENT_DATE)", connection))
                    {
                        command.Parameters.AddWithValue("@fullName", fullName);
                        command.Parameters.AddWithValue("@phoneNumber", phoneNumber);

                        command.ExecuteNonQuery();
                        MessageBox.Show("Клиент успешно добавлен!");
                        this.Hide();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при добавлении клиента: {ex.Message}");
                }
            }
        }
    }
}