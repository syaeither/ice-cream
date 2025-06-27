using Npgsql;
using System;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using DocumentFormat.OpenXml;
using System.Diagnostics;

namespace ice_cream
{
    public partial class EditOrderForm : Form
    {
        private int orderId;
        private int clientId;
        private string connectionString = Program.ConnectionString;

        public EditOrderForm(int orderId, int clientId)
        {
            InitializeComponent();
            this.orderId = orderId;
            this.clientId = clientId;
            LoadOrderData();
            LoadOrderProducts();
            LoadOrderHistory(clientId);
        }

        private void LoadOrderHistory(int clientId)
        {
            using (var connection = new NpgsqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    string query = @"
SELECT 
    o.date_orders AS ""Дата заказа"", 
    o.basket_products AS ""Товары"", 
    o.total_cost AS ""Общая сумма"", 
    e.full_name AS ""ФИО Сотрудника""
FROM 
    public.orders o
JOIN 
    public.employees e ON o.code_employee = e.id_employee
WHERE 
    o.code_client = @clientId
ORDER BY 
    o.date_orders DESC;";

                    using (var command = new NpgsqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("clientId", clientId);
                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                DataTable dt = new DataTable();
                                dt.Columns.Add("Дата заказа", typeof(DateTime));
                                dt.Columns.Add("Товары", typeof(string));
                                dt.Columns.Add("Общая сумма", typeof(decimal));
                                dt.Columns.Add("ФИО сотрудника", typeof(string));

                                while (reader.Read())
                                {
                                    DateTime orderDate = reader.GetDateTime(0);
                                    string basketProducts = reader.IsDBNull(1) ? "[]" : reader.GetString(1);
                                    decimal totalSum = reader.IsDBNull(2) ? 0 : reader.GetDecimal(2);
                                    string employeeName = reader.GetString(3);

                                    if (string.IsNullOrWhiteSpace(basketProducts))
                                    {
                                        basketProducts = "[]";
                                    }

                                    string productList;
                                    if (basketProducts.Trim().StartsWith("["))
                                    {
                                        try
                                        {
                                            var products = JsonConvert.DeserializeObject<List<Product>>(basketProducts);
                                            productList = string.Join(", ", products.Select(p => $"{p.Name} ({p.Quantity} шт.)"));
                                        }
                                        catch (JsonException)
                                        {
                                            productList = "Ошибка загрузки товаров";
                                        }
                                    }
                                    else
                                    {
                                        productList = basketProducts;
                                    }

                                    dt.Rows.Add(orderDate, productList, totalSum, employeeName);
                                }

                                dataGridView1.DataSource = dt;
                                dataGridView1.AllowUserToAddRows = false;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при загрузке истории заказов: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void LoadOrderProducts()
        {
            textBoxProducts.Clear();
            using (var connection = new NpgsqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (var command = new NpgsqlCommand("SELECT basket_products FROM public.orders WHERE id_order = @orderId", connection))
                    {
                        command.Parameters.AddWithValue("orderId", orderId);
                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read() && !reader.IsDBNull(0))
                            {
                                string basketProducts = reader.GetString(0);
                                string productList;

                                if (basketProducts.Trim().StartsWith("["))
                                {
                                    try
                                    {
                                        var products = JsonConvert.DeserializeObject<List<Product>>(basketProducts) ?? new List<Product>();
                                        productList = string.Join(Environment.NewLine, products.Select(p => $"{p.Name} | {p.Quantity} шт. | {p.Price:C}"));
                                    }
                                    catch (JsonException)
                                    {
                                        productList = "Ошибка загрузки товаров";
                                    }
                                }
                                else
                                {
                                    productList = basketProducts;
                                }

                                textBoxProducts.Text = productList;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    textBoxProducts.Text = "Ошибка загрузки товаров: " + ex.Message;
                }
            }
        }

        private void LoadOrderData()
        {
            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand(@"SELECT c.full_name, c.telephone, o.date_orders, o.completion_date, o.total_cost, o.status, o.payment_method 
                                                          FROM public.orders o JOIN public.clients c ON o.code_client = c.id_client 
                                                          WHERE o.id_order = @orderId", connection))
                {
                    command.Parameters.AddWithValue("orderId", orderId);
                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            textBoxClientFullName.Text = reader.GetString(0);
                            maskedPhone.Text = reader.GetString(1);
                            dateTimePickerOrderDate.Value = reader.GetDateTime(2);
                            dateTimePickerCompletionDate.Value = !reader.IsDBNull(3) ? reader.GetDateTime(3) : DateTime.Now;
                            numericUpDownTotalCost.Value = reader.GetDecimal(4);
                            comboBoxStatus.SelectedItem = reader.GetString(5);
                            comboBox1.SelectedItem = reader.IsDBNull(6) ? "Наличные" : reader.GetString(6);
                        }
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                string clientFullName = textBoxClientFullName.Text;
                string clientPhone = maskedPhone.Text;
                decimal totalCost = numericUpDownTotalCost.Value;
                string status = comboBoxStatus.SelectedItem?.ToString();
                DateTime completionDate = dateTimePickerCompletionDate.Value;
                string paymentMethod = comboBox1.SelectedItem?.ToString();

                var products = new List<Product>();
                foreach (string item in textBoxProducts.Lines)
                {
                    if (string.IsNullOrWhiteSpace(item)) continue; // Пропускаем пустые строки

                    var parts = item.Split('|');
                    if (parts.Length >= 1)
                    {
                        string name = parts[0].Trim();

                        //Если название товара пустое — пропускаем строку
                        if (string.IsNullOrWhiteSpace(name)) continue;

                        int quantity = 1; // Значение по умолчанию
                        decimal price = 0; // Значение по умолчанию

                        if (parts.Length >= 2)
                        {
                            int.TryParse(parts[1].Replace("шт.", "").Trim(), out quantity);
                        }
                        if (parts.Length >= 3)
                        {
                            decimal.TryParse(parts[2].Replace("₽", "").Trim(), out price);
                        }

                        products.Add(new Product
                        {
                            Name = name,
                            Quantity = quantity,
                            Price = price
                        });
                    }
                }

                string basketProducts = JsonConvert.SerializeObject(products);

                using (var connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();

                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            using (var command = new NpgsqlCommand(
                                "UPDATE public.clients SET full_name = @fullName, telephone = @telephone WHERE id_client = @clientId",
                                connection, transaction))
                            {
                                command.Parameters.AddWithValue("fullName", clientFullName);
                                command.Parameters.AddWithValue("telephone", clientPhone);
                                command.Parameters.AddWithValue("clientId", clientId);
                                command.ExecuteNonQuery();
                            }

                            using (var command = new NpgsqlCommand(
                                @"UPDATE public.orders 
                          SET status = @status, 
                              completion_date = @completionDate, 
                              payment_method = @paymentMethod, 
                              basket_products = @basketProducts,
                              total_cost = @totalCost
                          WHERE id_order = @orderId",
                                connection, transaction))
                            {
                                command.Parameters.AddWithValue("status", status);
                                command.Parameters.AddWithValue("completionDate", completionDate);
                                command.Parameters.AddWithValue("paymentMethod", paymentMethod);
                                command.Parameters.AddWithValue("basketProducts", basketProducts);
                                command.Parameters.AddWithValue("totalCost", totalCost);
                                command.Parameters.AddWithValue("orderId", orderId);
                                command.ExecuteNonQuery();
                            }

                            transaction.Commit();

                            MessageBox.Show("Данные успешно изменены!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            this.DialogResult = DialogResult.OK;
                            this.Close();
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            MessageBox.Show($"Ошибка при изменении данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить этот заказ?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                using (var connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    using (var command = new NpgsqlCommand("DELETE FROM public.orders WHERE id_order = @orderId", connection))
                    {
                        command.Parameters.AddWithValue("orderId", orderId);
                        command.ExecuteNonQuery();
                    }
                }
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            using (var connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand("UPDATE public.orders SET status = 'Завершен', completion_date = @completionDate WHERE id_order = @orderId", connection))
                {
                    command.Parameters.AddWithValue("completionDate", DateTime.Now);
                    command.Parameters.AddWithValue("orderId", orderId);
                    command.ExecuteNonQuery();
                }
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void EditOrderForm_Load(object sender, EventArgs e)
        {
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string clientName = textBoxClientFullName.Text;
            string clientPhone = maskedPhone.Text;
            DateTime orderDate = dateTimePickerOrderDate.Value;
            decimal totalCost = numericUpDownTotalCost.Value;
            string status = comboBoxStatus.SelectedItem?.ToString() ?? "Не указан";
            string paymentMethod = comboBox1.SelectedItem?.ToString() ?? "Не указан";

            var products = new List<Product>();
            foreach (string item in textBoxProducts.Lines)
            {
                var parts = item.Split('|');
                if (parts.Length >= 1) 
                {
                    products.Add(new Product
                    {
                        Name = parts[0].Trim()
                    });
                }
            }

            string orderNumber = GenerateOrderNumber();

            string fileName = $"Чек_Заказ_№_{orderNumber}.docx";
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                SectionProperties sectionProperties = new SectionProperties();
                body.AppendChild(sectionProperties);

                PageSize pageSize = new PageSize()
                {
                    Width = new UInt32Value((uint)ConvertMillimetersToTwips(139.7)),
                    Height = new UInt32Value((uint)ConvertMillimetersToTwips(215.9)) 
                };
                sectionProperties.AppendChild(pageSize);

                PageMargin pageMargin = new PageMargin()
                {
                    Top = new Int32Value(ConvertMillimetersToTwips(12.7)),
                    Right = new UInt32Value((uint)ConvertMillimetersToTwips(12.7)),
                    Bottom = new Int32Value(ConvertMillimetersToTwips(12.7)),
                    Left = new UInt32Value((uint)ConvertMillimetersToTwips(12.7)),
                    Header = new UInt32Value(0U),
                    Footer = new UInt32Value(0U),
                    Gutter = new UInt32Value(0U)
                };
                sectionProperties.AppendChild(pageMargin);

                Paragraph title = new Paragraph();
                Run runTitle = new Run();
                runTitle.AppendChild(new Text($"Товарный чек №{orderNumber}"));
                runTitle.RunProperties = new RunProperties(new Bold());
                runTitle.RunProperties.FontSize = new FontSize() { Val = "44" }; 
                title.AppendChild(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }));
                title.AppendChild(runTitle);
                body.AppendChild(title);

                AddParagraphWithLabel(body, "Клиент: ", clientName, fontSize: "24");  
                AddParagraphWithLabel(body, "Телефон: ", clientPhone, fontSize: "24"); 
                AddParagraphWithLabel(body, "Дата заказа: ", $"{orderDate:dd.MM.yyyy}", fontSize: "24");
                AddParagraphWithLabel(body, "Статус: ", status, fontSize: "24");
                AddParagraphWithLabel(body, "Способ оплаты: ", paymentMethod, fontSize: "24");

                AddParagraph(body, "--------------------------------------------------", fontSize: "24");

                AddParagraph(body, "Товары:", fontSize: "24", bold: true);
                foreach (var product in products)
                {
                    AddParagraph(body, $"{product.Name}", fontSize: "24");
                }

                AddParagraph(body, "--------------------------------------------------", fontSize: "24");

                AddParagraph(body, $"Общая сумма: {totalCost:C}", fontSize: "24", bold: true);

                Paragraph thankYouParagraph = new Paragraph();
                Run runThankYou = new Run();
                runThankYou.AppendChild(new Text("Спасибо за покупку!"));
                runThankYou.RunProperties = new RunProperties(new Bold());
                runThankYou.RunProperties.FontSize = new FontSize() { Val = "24" }; // 12pt (24 единицы)
                thankYouParagraph.AppendChild(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }));
                thankYouParagraph.AppendChild(runThankYou);
                body.AppendChild(thankYouParagraph);

                mainPart.Document.Save();
            }

            System.Diagnostics.Process.Start(filePath);
        }

        private int ConvertMillimetersToTwips(double millimeters)
        {
            return (int)(millimeters * 56.6929);
        }

        private void AddParagraphWithLabel(Body body, string label, string value, string fontSize = "24")
        {
            Paragraph paragraph = new Paragraph();

            Run runLabel = new Run();
            runLabel.AppendChild(new Text(label + " "));
            runLabel.RunProperties = new RunProperties(new Bold());
            runLabel.RunProperties.FontSize = new FontSize() { Val = fontSize };
            paragraph.AppendChild(runLabel);

            Run runValue = new Run();
            runValue.AppendChild(new Text(value));
            runValue.RunProperties = new RunProperties();
            runValue.RunProperties.FontSize = new FontSize() { Val = fontSize };
            paragraph.AppendChild(runValue);

            body.AppendChild(paragraph);
        }

        private void AddParagraph(Body body, string text, bool bold = false, string fontSize = "24")
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            run.AppendChild(new Text(text));
            run.RunProperties = new RunProperties();
            if (bold)
            {
                run.RunProperties.Bold = new Bold();
            }
            run.RunProperties.FontSize = new FontSize() { Val = fontSize };
            paragraph.AppendChild(run);
            body.AppendChild(paragraph);
        }

        private string GenerateOrderNumber()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        public class Product
        {
            public string Name { get; set; }
            public int Quantity { get; set; }
            public decimal Price { get; set; }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            programinfo form2 = new programinfo();
            form2.Show();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string helpFileName = "Справка формы редактирования заказа.pdf";

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