using System;
using System.Diagnostics;
using System.Windows.Forms;
using Npgsql;
using System.IO;

namespace ice_cream
{
    public partial class Authorization : Form
    {
        private bool isPasswordVisible = false;

        public Authorization()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) //кнопка Войти
        {
            string username = textBox1.Text.Trim();
            string password = textBox2.Text.Trim();

            using (NpgsqlConnection connection = new NpgsqlConnection(Program.ConnectionString))
            {
                connection.Open();
                string query = "SELECT id_employee, full_name, position, password FROM employees WHERE username = @Username";
                using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Username", username);
                    using (NpgsqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            int employeeId = reader.GetInt32(0);
                            string fullName = reader.GetString(1);
                            string role = reader.GetString(2);
                            string storedPassword = reader.GetString(3);

                            if ((password) == storedPassword)
                            {
                                UserSession.EmployeeId = employeeId;
                                UserSession.FullName = fullName;
                                UserSession.Position = role;
                                UserSession.Username = username;

                                programinfo infoForm = new programinfo();

                                switch (role)
                                {
                                    case "Администратор":
                                        administrator adminForm = new administrator();
                                        adminForm.Show();
                                        this.Hide();
                                        break;
                                    case "Руководитель":
                                        supervisor managerForm = new supervisor();
                                        managerForm.Show();
                                        this.Hide();
                                        break;
                                    case "Продавец":
                                        salesman sellerForm = new salesman(employeeId);
                                        sellerForm.Show();
                                        this.Hide();
                                        break;
                                    default:
                                        MessageBox.Show("Неверная роль пользователя.");
                                        break;
                                }
                            }
                            else
                            {
                                MessageBox.Show("Неправильный пароль.");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Неверные учётные данные. Пожалуйста, повторите попытку.");
                        }
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e) //кнопка глаз
        {
            if(isPasswordVisible)
            {
                textBox2.PasswordChar = '*';
                button2.BackgroundImage = Properties.Resources.закрытый_глаз;
                button2.BackgroundImageLayout = ImageLayout.Zoom; ;
            }
            else
            {
                textBox2.PasswordChar = '\0';
                button2.BackgroundImage = Properties.Resources.открытый_глаз;
                button2.BackgroundImageLayout = ImageLayout.Zoom;
            }
            isPasswordVisible = !isPasswordVisible;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) //ссылка на восстановление пароля
        {
            PasswordRecoveryForm recoveryForm = new PasswordRecoveryForm();
            recoveryForm.ShowDialog();
        }

        private void pictureBox2_Click_1(object sender, EventArgs e) //кнопка О программе
        {
            programinfo form2 = new programinfo();
            form2.ShowDialog();
        }

        private void pictureBox3_Click(object sender, EventArgs e) //кнопка Справка
        {
            string helpFileName = "Справка формы авторизации.pdf";

            string appDirectory = AppDomain.CurrentDomain.BaseDirectory;

            string helpFilesPath = Path.Combine(appDirectory, "HelpFiles");

            string fullFilePath = Path.Combine(helpFilesPath, helpFileName);

            try
            {
                if (File.Exists(fullFilePath))
                {
                    Process.Start(new ProcessStartInfo(fullFilePath) { UseShellExecute = true });
                }
                else
                {
                    MessageBox.Show($"Файл справки не найден: {helpFileName}\nОжидаемый путь: {helpFilesPath}",
                                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось открыть файл справки: {ex.Message}",
                                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}