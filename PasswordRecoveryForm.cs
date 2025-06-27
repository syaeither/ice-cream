using System;
using System.Net;
using System.Net.Mail;
using System.Windows.Forms;
using Npgsql;
using System.Security.Cryptography;
using System.Text;
using System.Diagnostics;
using System.IO;

namespace ice_cream
{
    public partial class PasswordRecoveryForm : Form
    {
        private string connectionString = Program.ConnectionString;
        private string verificationCode;
        private string userEmail;
        private int employeeId;
        private bool isPasswordVisible = false;

        private const string smtpHost = "smtp.gmail.com";
        private const int smtpPort = 587;
        private const string smtpUsername = "alexeevaalexandra26@gmail.com";
        private const string smtpPassword = "jbof ufcf xspe vstr";

        public PasswordRecoveryForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            userEmail = textBox1.Text.Trim();

            if (string.IsNullOrEmpty(userEmail))
            {
                MessageBox.Show("Введи свою почту, брат", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!IsEmailValid(userEmail))
            {
                MessageBox.Show("Это не похоже на нормальный email", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                if (!GetEmployeeByEmail(userEmail))
                {
                    MessageBox.Show("Такой email не зарегистрирован", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                verificationCode = GenerateVerificationCode();
                SendVerificationEmail(userEmail, verificationCode);

                MessageBox.Show("Код подтверждения полетел на твою почту", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Бро, ошибка: {ex.Message}", "Критическая ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string enteredCode = textBox2.Text.Trim();

            if (string.IsNullOrEmpty(enteredCode))
            {
                MessageBox.Show("Введи код, который пришел на почту", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (enteredCode == verificationCode)
            {
                MessageBox.Show("Код верный, молодец!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox3.Text = "";
            }
            else
            {
                MessageBox.Show("Неверный код, попробуй еще раз", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string newPassword = textBox3.Text;

            if (string.IsNullOrEmpty(newPassword))
            {
                MessageBox.Show("Введите новый пароль", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (newPassword.Length < 6)
            {
                MessageBox.Show("Пароль слишком короткий (минимум 6 символов)", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string hashedPassword = HashPassword(newPassword);
                UpdateEmployeePassword(employeeId, newPassword);

                MessageBox.Show("Пароль успешно изменен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при смене пароля: {ex.Message}", "Критическая ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private bool IsEmailValid(string email)
        {
            try
            {
                var addr = new MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        private bool GetEmployeeByEmail(string email)
        {
            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var cmd = new NpgsqlCommand("SELECT id_employee FROM employees WHERE email = @email", connection))
                {
                    cmd.Parameters.AddWithValue("@email", email);
                    var result = cmd.ExecuteScalar();
                    if (result != null && result != DBNull.Value)
                    {
                        employeeId = Convert.ToInt32(result);
                        return true;
                    }
                }
            }
            return false;
        }

        private string GenerateVerificationCode()
        {
            Random random = new Random();
            return random.Next(100000, 999999).ToString();
        }

        private void SendVerificationEmail(string email, string code)
        {
            using (SmtpClient client = new SmtpClient(smtpHost, smtpPort))
            {
                client.EnableSsl = true;
                client.Credentials = new NetworkCredential(smtpUsername, smtpPassword);

                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(smtpUsername);
                mailMessage.To.Add(email);
                mailMessage.Subject = "Восстановление пароля - Ice Cream System";
                mailMessage.Body = $"Твой код подтверждения: {code}\n\n" +
                                 $"Введи этот код в программе для восстановления пароля.\n" +
                                 $"Код действителен 10 минут.\n\n" +
                                 $"Если это не ты запрашивал смену пароля - срочно напиши в поддержку!";

                client.Send(mailMessage);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (isPasswordVisible)
            {
                textBox3.PasswordChar = '*';
                button5.BackgroundImage = Properties.Resources.закрытый_глаз;
                button5.BackgroundImageLayout = ImageLayout.Zoom; ;
            }
            else
            {
                textBox3.PasswordChar = '\0';
                button5.BackgroundImage = Properties.Resources.открытый_глаз;
                button5.BackgroundImageLayout = ImageLayout.Zoom;
            }
            isPasswordVisible = !isPasswordVisible;
        }

        private string HashPassword(string password)
        {
            using (SHA1 sha1 = SHA1.Create())
            {
                byte[] hashedBytes = sha1.ComputeHash(Encoding.UTF8.GetBytes(password));
                return BitConverter.ToString(hashedBytes).Replace("-", "").ToLower();
            }
        }

        private void UpdateEmployeePassword(int id, string password)
        {
            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var cmd = new NpgsqlCommand("UPDATE employees SET password = @password WHERE id_employee = @id", connection))
                {
                    cmd.Parameters.AddWithValue("@password", password);
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string helpFileName = "Справка по форме восстановление пароля.pdf";

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