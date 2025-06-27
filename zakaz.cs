using System;
using Npgsql;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace ice_cream
{
    public partial class zakaz : Form
    {
        private decimal totalAmount = 0;
        private decimal discount = 0;
        private Label labelDiscount;

        public zakaz()
        {
            InitializeComponent();
            InitializeCustomComponents();
        }

        private void InitializeCustomComponents()
        {
            labelDiscount = new Label { Text = "0%", Location = new System.Drawing.Point(713, 428), Size = new System.Drawing.Size(61, 19), Font = new System.Drawing.Font("Times New Roman", 12)};
            this.Controls.Add(labelDiscount);
            LoadCategories();
        }

        private void LoadCategories()
        {
            using (var connection = new NpgsqlConnection(Program.ConnectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand("SELECT DISTINCT categories FROM products", connection))
                {
                    using (var reader = command.ExecuteReader())
                    {
                        comboBox1.Items.Clear();
                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(0))
                            {
                                comboBox1.Items.Add(reader.GetString(0));
                            }
                        }
                    }
                }
            }
        }

        private void CheckClient(string phoneNumber)
        {
            using (var connection = new NpgsqlConnection(Program.ConnectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand("SELECT full_name, id_client FROM clients WHERE telephone = @phoneNumber", connection))
                {
                    command.Parameters.AddWithValue("@phoneNumber", phoneNumber);

                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string fullName = reader.GetString(0);
                            int clientId = reader.GetInt32(1); 

                            MessageBox.Show($"Клиент найден: {fullName}");
                         
                            if (clientId != 0)
                            {
                                discount = 20;
                                labelDiscount.Text = $"{discount}%";
                            }
                            else
                            {
                                discount = 0;
                                labelDiscount.Text = "0%";
                            }

                            UpdateTotalAmount(); 
                        }
                        else
                        {
                            DialogResult dialogResult = MessageBox.Show("Клиент не найден. Добавить клиента?", "Вопрос", MessageBoxButtons.YesNo);
                            if (dialogResult == DialogResult.Yes)
                            {
                                ShowAddClientForm();
                            }
                        }
                    }
                }
            }
        }

        private void LoadProductsByCategory(string category)
        {
            using (var connection = new NpgsqlConnection(Program.ConnectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand("SELECT id_products, name, price, stock_quantity FROM products WHERE categories = @category", connection))
                {
                    command.Parameters.AddWithValue("@category", category);

                    using (var reader = command.ExecuteReader())
                    {
                        listBox1.Items.Clear();
                        while (reader.Read())
                        {
                            listBox1.Items.Add(new Product
                            {
                                Id = reader.GetInt32(0),
                                Name = reader.GetString(1),
                                Price = reader.GetDecimal(2),
                                StockQuantity = reader.GetInt32(3)
                            });
                        }
                    }
                }
            }
        }

        private void AddToCart(int quantity)
        {
            if (listBox1.SelectedItem is Product selectedProduct)
            {             
                listBox2.Items.Add(new CartItem
                {
                    Product = selectedProduct,
                    Quantity = quantity
                });

                totalAmount += selectedProduct.Price * quantity;
                UpdateTotalAmount();
            }
        }

        private void UpdateTotalAmount()
        {
            decimal discountedAmount = totalAmount - (totalAmount * discount / 100);
            textBox1.Text = discountedAmount.ToString("C2");
        }

        private void ShowAddClientForm()
        {
            AddClientForm addClientForm = new AddClientForm();
            addClientForm.ShowDialog();
        }


        private void button4_Click(object sender, EventArgs e)
        {
            if (listBox2.Items.Count == 0)
            {
                MessageBox.Show("Корзина пуста. Добавьте товары перед оформлением заказа.");
                return;
            }

            string phoneNumber = maskedTextBox1.Text;
            int? clientId = GetClientIdByPhoneNumber(phoneNumber);

            if (clientId == null || clientId == -1)
            {
                MessageBox.Show("Клиент не найден. Заказ будет оформлен без привязки к клиенту.");
                clientId = null;
                discount = 0;
                labelDiscount.Text = "0%";
            }
            else if (clientId == 0)
            {
                discount = 0;
                labelDiscount.Text = "0%";
            }

            decimal totalAmount = CalculateTotalAmountWithDiscount();
            int totalQuantity = listBox2.Items.Cast<CartItem>().Sum(item => item.Quantity);

            string paymentMethod = comboBoxPaymentMethod.SelectedItem?.ToString();

            if (string.IsNullOrEmpty(paymentMethod))
            {
                MessageBox.Show("Выберите способ оплаты.");
                return;
            }

            // Получаем ID текущего сотрудника из сессии
            int? employeeId = UserSession.EmployeeId;

            using (var connection = new NpgsqlConnection(Program.ConnectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        var productNames = listBox2.Items.Cast<CartItem>()
                            .Select(item => item.Product.Name)
                            .ToList();

                        string productNamesString = string.Join(", ", productNames);

                        using (var orderCommand = new NpgsqlCommand("INSERT INTO orders (code_client, date_orders, total_cost," +
                            "status, payment_method, basket_products, code_employee) VALUES (@clientId, @orderDate, @totalAmount, @status, @paymentMethod," +
                            " @basketProducts, @employeeId) RETURNING id_order", connection, transaction))
                        {
                            orderCommand.Parameters.AddWithValue("@clientId", (object)clientId ?? DBNull.Value);
                            orderCommand.Parameters.AddWithValue("@orderDate", DateTime.Now);
                            orderCommand.Parameters.AddWithValue("@totalAmount", totalAmount);
                            orderCommand.Parameters.AddWithValue("@status", "В процессе");
                            orderCommand.Parameters.AddWithValue("@paymentMethod", paymentMethod);
                            orderCommand.Parameters.AddWithValue("@basketProducts", productNamesString);
                            orderCommand.Parameters.AddWithValue("@employeeId", (object)employeeId ?? DBNull.Value); // Добавляем ID сотрудника

                            int orderId = (int)orderCommand.ExecuteScalar();

                            foreach (CartItem item in listBox2.Items)
                            {
                                using (var updateStockCommand = new NpgsqlCommand("UPDATE products SET stock_quantity = " +
                                    "stock_quantity - @quantity WHERE id_products = @productId", connection, transaction))
                                {
                                    updateStockCommand.Parameters.AddWithValue("@quantity", item.Quantity);
                                    updateStockCommand.Parameters.AddWithValue("@productId", item.Product.Id);
                                    updateStockCommand.ExecuteNonQuery();
                                }
                            }

                            transaction.Commit();
                            MessageBox.Show("Заказ успешно оформлен!");
                        }
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        MessageBox.Show($"Ошибка при оформлении заказа: {ex.Message}");
                    }
                }
            }
        }

        private int GetClientIdByPhoneNumber(string phoneNumber)
        {
            using (var connection = new NpgsqlConnection(Program.ConnectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand("SELECT id_client FROM clients WHERE telephone = @phoneNumber", connection))
                {
                    command.Parameters.AddWithValue("@phoneNumber", phoneNumber);

                    var result = command.ExecuteScalar();
                    return result != null ? (int)result : -1;
                }
            }
        }

        private decimal CalculateTotalAmountWithDiscount()
        {
            decimal totalAmount = 0;

            foreach (CartItem item in listBox2.Items)
            {
                decimal itemTotal = item.Product.Price * item.Quantity;

                decimal discount = 0.10m; 
                itemTotal -= itemTotal * discount;

                totalAmount += itemTotal;
            }
            return totalAmount;
        }

        private void ClearCart()
        {
            listBox2.Items.Clear();
            totalAmount = 0;
            discount = 0;
            labelDiscount.Text = "0%";
            textBox1.Text = "0"; 
            maskedTextBox1.Clear(); 
            listBox1.Items.Clear(); 
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedItem is CartItem selectedItem)
            {
                totalAmount -= selectedItem.Product.Price * selectedItem.Quantity;
                listBox2.Items.Remove(selectedItem);
                UpdateTotalAmount(); 
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {       
            ClearCart();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int quantity = 1;
            AddToCart(quantity);
        }

        private void button3_Click(object sender, EventArgs e)
        {      
            if (comboBox1.SelectedItem != null)
            {
                LoadProductsByCategory(comboBox1.SelectedItem.ToString());
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите категорию.");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CheckClient(maskedTextBox1.Text);
        }

        public class Product
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public decimal Price { get; set; }
            public int StockQuantity { get; set; }

            public override string ToString()
            {
                return $"{Name} - {Price:C} (в наличии: {StockQuantity})";
            }
        }

        public class CartItem
        {
            public Product Product { get; set; }
            public int Quantity { get; set; }

            public override string ToString()
            {
                return $"{Product.Name} - {Quantity} шт. - {Product.Price * Quantity:C}";
            }
        }

        private int GenerateSaleId()
        {
            return new Random().Next(1000, 9999);
        }    

        private void button8_Click_1(object sender, EventArgs e)
        {
            //чек
            GenerateReceipt();
        }

        private void GenerateReceipt()
        {
            if (listBox2.Items.Count == 0)
            {
                MessageBox.Show("Корзина пуста. Оформите заказ перед печатью чека.");
                return;
            }

            string receiptNumber = GenerateReceiptNumber();
            string storeName = "Магазин мороженного Ice-Cream";
            string storeAddress = "г. Ростов-на-Дону, ул. Ледяная, д. 5";
            string dateTime = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
            string paymentMethod = comboBoxPaymentMethod.SelectedItem?.ToString() ?? "Не указано";
            decimal totalAmount = CalculateTotalAmountWithDiscount();

            string filePath = $"Чек_{receiptNumber}.docx";

            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                SectionProperties sectionProps = new SectionProperties();
                PageSize pageSize = new PageSize()
                {
                    Width = new UInt32Value((uint)ConvertMillimetersToTwips(148)),
                    Height = new UInt32Value((uint)ConvertMillimetersToTwips(210))
                };
                sectionProps.Append(pageSize);

                AddParagraph(body, storeName, true, 22, JustificationValues.Center);
                AddParagraph(body, storeAddress, false, 16, JustificationValues.Center);
                AddParagraph(body, "", false, 12);

                AddParagraph(body, $"Товарный чек № {receiptNumber}", true, 16);
                AddParagraph(body, $"Дата: {dateTime}", false, 14);
                AddParagraph(body, new string('_', 50), false, 12);

                AddParagraph(body, "Список товаров:", true, 16);
                foreach (CartItem item in listBox2.Items)
                {
                    AddParagraph(body, $"{item.Product.Name} x {item.Quantity} - {item.Product.Price * item.Quantity:C}", false, 12);
                }

                AddParagraph(body, new string('_', 50), false, 12);
                AddParagraph(body, $"Оплата: {paymentMethod}", false, 14);
                AddParagraph(body, $"Итого: {totalAmount:C}", true, 16);
                AddParagraph(body, "Спасибо за покупку!", true, 14, JustificationValues.Center);

                doc.MainDocumentPart.Document.Save();
            }

            MessageBox.Show("Чек сохранен в файле: " + filePath);
            System.Diagnostics.Process.Start(filePath);
        }

        private int ConvertMillimetersToTwips(double millimeters)
        {
            return (int)(millimeters * 56.6929);
        }

        private void AddParagraph(Body body, string text, bool bold, int fontSize, JustificationValues? justification = null)
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            RunProperties runProperties = new RunProperties();

            if (bold) runProperties.Bold = new Bold();
            runProperties.FontSize = new FontSize() { Val = (fontSize * 2).ToString() };

            run.Append(runProperties);
            run.Append(new Text(text));
            paragraph.Append(run);

            if (justification.HasValue)
            {
                paragraph.ParagraphProperties = new ParagraphProperties(
                    new Justification() { Val = justification.Value });
            }
            body.Append(paragraph);
        }

        private string GenerateReceiptNumber()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            administrator adminForm = new administrator();
            adminForm.Show();
            this.Hide();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                string phoneNumber = maskedTextBox1.Text.Trim();

                if (string.IsNullOrEmpty(phoneNumber))
                {
                    MessageBox.Show("Введите номер телефона клиента", "Ошибка",
                                 MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int? clientId = GetClientIdByPhoneNumber(phoneNumber);

                if (clientId == null || clientId == -1)
                {
                    MessageBox.Show("Клиент не найден в базе данных", "Клиент не найден",
                                 MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                bool hasHistory = CheckClientHistory((int)clientId);

                if (!hasHistory)
                {
                    MessageBox.Show("У данного клиента нет истории заказов", "Нет истории",
                                 MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                history_client historyForm = new history_client((int)clientId);
                historyForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool CheckClientHistory(int clientId)
        {
            using (var connection = new NpgsqlConnection(Program.ConnectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand("SELECT COUNT(*) FROM orders WHERE code_client = @clientId", connection))
                {
                    command.Parameters.AddWithValue("@clientId", clientId);
                    long count = (long)command.ExecuteScalar();
                    return count > 0;
                }
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            programinfo form2 = new programinfo();
            form2.Show();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string helpFileName = "Справка формы оформление заказа.pdf";

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