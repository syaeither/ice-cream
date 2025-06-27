using Npgsql;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace ice_cream
{
    public partial class administrator : Form
    {
        private int? selectedOrderId = null;
        private int orderId;     
        string connectionString = Program.ConnectionString;

        public administrator()
        {
            InitializeComponent();
            LoadAndDisplayOrders();
        }

        private void LoadAndDisplayOrders(string searchTerm = "") //Список заказов
        {
            panel1.Controls.Clear();

            string query = @"
            SELECT 
                o.id_order, 
                COALESCE(c.full_name, 'Без клиента') AS full_name, 
                COALESCE(c.telephone, '') AS telephone, 
                o.date_orders, 
                o.completion_date, 
                o.total_cost, 
                o.status, 
                o.payment_method,
                o.basket_products,
                COALESCE(o.code_client, -1) AS code_client
            FROM 
                public.orders o
            LEFT JOIN 
                public.clients c ON o.code_client = c.id_client
            WHERE 
                1=1";

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                query += " AND (COALESCE(c.full_name, 'Без клиента') ILIKE @searchTerm OR o.status ILIKE @searchTerm OR COALESCE(c.telephone, '') ILIKE @searchTerm)";
            }

            query += " ORDER BY o.date_orders DESC;";

            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(query, connection))
                {
                    if (!string.IsNullOrWhiteSpace(searchTerm))
                    {
                        command.Parameters.AddWithValue("searchTerm", "%" + searchTerm + "%");
                    }

                    using (var reader = command.ExecuteReader())
                    {
                        int yOffset = 10;

                        while (reader.Read())
                        {
                            int orderId = reader.GetInt32(0);
                            string clientName = reader.GetString(1);
                            string clientPhone = reader.GetString(2);
                            DateTime dateOrders = reader.GetDateTime(3);
                            DateTime? completionDate = reader.IsDBNull(4) ? (DateTime?)null : reader.GetDateTime(4);
                            decimal totalCost = reader.GetDecimal(5);
                            string status = reader.GetString(6);
                            string paymentMethod = reader.IsDBNull(7) ? "Не указан" : reader.GetString(7);
                            string basketProducts = reader.IsDBNull(8) ? "[]" : reader.GetString(8);
                            int clientId = reader.GetInt32(9);

                            Panel orderPanel = new Panel
                            {
                                Width = 670,
                                Height = 80,
                                Location = new Point(10, yOffset),
                                BackColor = Color.White,
                                BorderStyle = BorderStyle.FixedSingle,
                                Tag = new { OrderId = orderId, ClientId = clientId, Status = status }
                            };

                            if (status == "В процессе")
                                orderPanel.BackColor = Color.LightBlue;
                            else if (status == "Завершен")
                                orderPanel.BackColor = Color.LightGreen;

                            orderPanel.Controls.Add(new Label
                            {
                                Text = $"{clientName} / {clientPhone}",
                                Location = new Point(6, 10),
                                Font = new Font("Times New Roman", 12.25F),
                                AutoSize = true
                            });

                            orderPanel.Controls.Add(new Label
                            {
                                Text = $"{totalCost:C}",
                                Location = new Point(orderPanel.Width - 70, 10),
                                AutoSize = true
                            });

                            orderPanel.Controls.Add(new Label
                            {
                                Text = $"Оплата: {paymentMethod}",
                                Location = new Point(6, 30),
                                AutoSize = true
                            });

                            orderPanel.Controls.Add(new Label
                            {
                                Text = $"Статус: {status}",
                                Location = new Point(6, 50),
                                AutoSize = true
                            });

                            orderPanel.DoubleClick += (s, e) =>
                            {
                                var tag = (dynamic)((Panel)s).Tag;
                                string orderStatus = tag.Status;

                                if (orderStatus == "Завершен")
                                {
                                    MessageBox.Show("Этот заказ уже завершен и не может быть отредактирован.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    return;
                                }

                                int selectedOrderId = tag.OrderId;
                                int selectedClientId = tag.ClientId;

                                EditOrderForm editForm = new EditOrderForm(selectedOrderId, selectedClientId);
                                editForm.ShowDialog();

                                LoadAndDisplayOrders();
                            };

                            panel1.Controls.Add(orderPanel);
                            yOffset += orderPanel.Height + 5;
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e) //кнопка назад к авторизацию
        {
            Authorization form2 = new Authorization();
            form2.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e) //кнопка к переходу к форме клиенты
        {
            clients form = new clients();
            form.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e) //кнопка к переходу к форме сотрудники
        {
            employees form2 = new employees();
            form2.Show();
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e) //кнопка к переходу к форме товары
        {
            products form2 = new products();
            form2.Show();
            this.Hide();
        }

        private void button7_Click(object sender, EventArgs e) //кнопка поиска
        {
            string searchTerm = textBox1.Text.Trim();
            LoadAndDisplayOrders(searchTerm);
        }

        private void button5_Click(object sender, EventArgs e) //кнопка к переходу к форме оформление заказ
        {
            zakaz form2 = new zakaz(); //форма оформление заказа
            form2.Show();
            this.Hide();
        }

        private void button6_Click(object sender, EventArgs e) //кнопка к переходу к форме поставок
        {
            supplier_shipments form2 = new supplier_shipments();
            form2.Show();
            this.Hide();
        }

        private void pictureBox2_Click(object sender, EventArgs e) //кнопка О программе
        {
            programinfo form2 = new programinfo();
            form2.Show();
        }

        private void pictureBox3_Click(object sender, EventArgs e) //кнопка Справка
        {
            string helpFileName = "Справка формы администратора.pdf";

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