using Npgsql;
using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ice_cream
{
    public partial class salesman : Form
    {
        private decimal totalAmount = 0;
        private int codeEmployee;

        public salesman(int employeeId)
        {
            InitializeComponent();
            codeEmployee = employeeId;
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
                if (selectedProduct.StockQuantity < quantity)
                {
                    MessageBox.Show("Недостаточно товара на складе.");
                    return;
                }

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
            if (textBox1 != null)
            {
                textBox1.Text = totalAmount.ToString("C2");
            }
            else
            {
                MessageBox.Show("textBox1 не инициализирован.");
            }
        }

        private void ClearCart()
        {
            listBox2.Items.Clear();
            totalAmount = 0;
            textBox1.Text = "0";
            listBox1.Items.Clear();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Authorization form2 = new Authorization();
            form2.Show();
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (listBox2.Items.Count == 0)
            {
                MessageBox.Show("Корзина пуста. Добавьте товары перед оформлением заказа.");
                return;
            }

            if (comboBoxPaymentMethod.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, выберите способ оплаты.");
                return;
            }

            using (var connection = new NpgsqlConnection(Program.ConnectionString))
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        int saleId;
                        using (var saleCommand = new NpgsqlCommand("INSERT INTO sales (code_employee, order_date, total_amount, payment_method) VALUES (@codeEmployee, @orderDate, @totalAmount, @paymentMethod) RETURNING id_sale", connection, transaction))
                        {
                            saleCommand.Parameters.AddWithValue("@codeEmployee", codeEmployee);
                            saleCommand.Parameters.AddWithValue("@orderDate", DateTime.Now);
                            saleCommand.Parameters.AddWithValue("@totalAmount", totalAmount);
                            saleCommand.Parameters.AddWithValue("@paymentMethod", comboBoxPaymentMethod.SelectedItem.ToString());
                            saleId = (int)saleCommand.ExecuteScalar();
                        }

                        foreach (CartItem item in listBox2.Items)
                        {
                            using (var basketCommand = new NpgsqlCommand("INSERT INTO basket (code_sales, code_product, quantity, price) VALUES (@codeSales, @codeProduct, @quantity, @price)", connection, transaction))
                            {
                                basketCommand.Parameters.AddWithValue("@codeSales", saleId);
                                basketCommand.Parameters.AddWithValue("@codeProduct", item.Product.Id);
                                basketCommand.Parameters.AddWithValue("@quantity", item.Quantity);
                                basketCommand.Parameters.AddWithValue("@price", item.Product.Price);
                                basketCommand.ExecuteNonQuery();
                            }

                            using (var updateStockCommand = new NpgsqlCommand("UPDATE products SET stock_quantity = stock_quantity - @quantity WHERE id_products = @codeProduct", connection, transaction))
                            {
                                updateStockCommand.Parameters.AddWithValue("@quantity", item.Quantity);
                                updateStockCommand.Parameters.AddWithValue("@codeProduct", item.Product.Id);
                                updateStockCommand.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                        MessageBox.Show("Продажа успешно оформлен!");

                        DialogResult result = MessageBox.Show("Хотите открыть чек?", "Чек", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            CreateCheck(saleId);
                        }

                        ClearCart();
                    }
                    catch (Exception ex)
                    {
                        if (transaction.Connection != null)
                        {
                            try
                            {
                                transaction.Rollback();
                            }
                            catch (Exception rollbackEx)
                            {
                                MessageBox.Show($"Ошибка отката транзакции: {rollbackEx.Message}");
                            }
                        }
                        MessageBox.Show($"Ошибка при оформлении заказа: {ex.Message}\n{ex.StackTrace}");
                    }
                }
            }
        }

        private void CreateCheck(int saleId)
        {
            string fileName = $"Check_{saleId}.docx";
            string filePath = System.IO.Path.Combine(Environment.CurrentDirectory, fileName);

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                string sellerName = GetSellerName(codeEmployee);
                string paymentMethod = comboBoxPaymentMethod.SelectedItem.ToString();

                body.Append(CreateParagraph($"Чек № {saleId}", true, 14));
                body.Append(CreateParagraph("Наименование магазина: Магазин мороженого «Ice-Cream»", false));
                body.Append(CreateParagraph("Адрес: г. Ростов-на-Дону, Большая Садовая ул., 78", false));
                body.Append(CreateParagraph($"Продавец: {sellerName}", false));
                body.Append(CreateParagraph($"Способ оплаты: {paymentMethod}", false));

                Table table = new Table();

                TableProperties tblProps = new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                        new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                        new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                        new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                        new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                        new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 }
                    ),
                    new TableWidth() { Width = "100%", Type = TableWidthUnitValues.Pct }
                );
                table.AppendChild(tblProps);

                TableGrid grid = new TableGrid();
                grid.Append(new GridColumn() { Width = "10%" });  
                grid.Append(new GridColumn() { Width = "40%" }); 
                grid.Append(new GridColumn() { Width = "20%" }); 
                grid.Append(new GridColumn() { Width = "15%" }); 
                grid.Append(new GridColumn() { Width = "15%" });
                table.Append(grid);

                TableRow headerRow = new TableRow();
                headerRow.Append(CreateTableCell("№", true, JustificationValues.Center));
                headerRow.Append(CreateTableCell("Наименование товара", true, JustificationValues.Center));
                headerRow.Append(CreateTableCell("Цена за 1 ед.", true, JustificationValues.Center));
                headerRow.Append(CreateTableCell("Количество", true, JustificationValues.Center));
                headerRow.Append(CreateTableCell("Сумма", true, JustificationValues.Center));
                table.Append(headerRow);

                int index = 1;
                foreach (CartItem item in listBox2.Items)
                {
                    TableRow row = new TableRow();
                    row.Append(CreateTableCell(index.ToString(), false, JustificationValues.Center));
                    row.Append(CreateTableCell(item.Product.Name, false, JustificationValues.Left));
                    row.Append(CreateTableCell(item.Product.Price.ToString("C2"), false, JustificationValues.Center));
                    row.Append(CreateTableCell(item.Quantity.ToString(), false, JustificationValues.Center));
                    row.Append(CreateTableCell((item.Product.Price * item.Quantity).ToString("C2"), false, JustificationValues.Center));
                    table.Append(row);
                    index++;
                }

                body.Append(table);
                body.Append(CreateParagraph($"Всего наименований: {listBox2.Items.Count}, на сумму {totalAmount.ToString("C2")}", false));
                body.Append(CreateParagraph($"Итого: {totalAmount.ToString("C2")}", true, 14));
                mainPart.Document.Save();
            }
            System.Threading.Thread.Sleep(1000);
            Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
        }

        private TableCell CreateTableCell(string text, bool isBold = false, JustificationValues? alignment = null)
        {
            if (alignment == null) alignment = JustificationValues.Left; 

            RunProperties runProps = new RunProperties();
            runProps.Append(new FontSize() { Val = "24" }); 
            if (isBold) runProps.Append(new Bold());

            Run run = new Run(runProps, new Text(text));
            Paragraph paragraph = new Paragraph(new ParagraphProperties(new Justification() { Val = alignment.Value }), run);
            TableCell cell = new TableCell(paragraph);

            return cell;
        }

        private Paragraph CreateParagraph(string text, bool isBold = false, int fontSize = 12)
        {
            RunProperties runProps = new RunProperties();
            runProps.Append(new FontSize() { Val = (fontSize * 2).ToString() });
            if (isBold) runProps.Append(new Bold());

            Run run = new Run();
            run.Append(runProps);
            run.Append(new Text(text));

            Paragraph paragraph = new Paragraph(run);
            return paragraph;
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

        private string GetSellerName(int employeeId)
        {
            using (var connection = new NpgsqlConnection(Program.ConnectionString))
            {
                connection.Open();
                using (var command = new NpgsqlCommand("SELECT full_name FROM employees WHERE id_employee = @employeeId", connection))
                {
                    command.Parameters.AddWithValue("@employeeId", employeeId);
                    return command.ExecuteScalar()?.ToString() ?? "Неизвестный продавец";
                }
            }
        }

        private int GetLastSaleId()
        {
            using (var connection = new NpgsqlConnection("Host=localhost;Port=5432;Username=postgres;Password=1234;Database=ice_cream"))
            {
                connection.Open();
                using (var command = new NpgsqlCommand("SELECT MAX(id_sale) FROM sales", connection))
                {
                    return (int)(command.ExecuteScalar() ?? 0);
                }
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            programinfo form2 = new programinfo();
            form2.Show();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            string helpFileName = "Справка формы продавца.pdf";

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