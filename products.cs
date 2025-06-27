using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;

namespace ice_cream
{
    public partial class products : Form
    {
        string connectionString = Program.ConnectionString;
        private int selectedProductId;

        public products()
        {
            InitializeComponent();
            LoadData();
            dataGridView1.AllowUserToAddRows = false;
            textBox2.KeyPress += textBox2_KeyPress;
        }

        private void LoadData()
        {
            using (var conn = new NpgsqlConnection(connectionString))
            {
                conn.Open();
                var adapter = new NpgsqlDataAdapter("SELECT * FROM products", conn);
                var table = new DataTable();
                adapter.Fill(table);

                dataGridView1.DataSource = table;
                dataGridView1.Columns["id_products"].Visible = false; ;
                dataGridView1.Columns["name"].HeaderText = "Наименование";
                dataGridView1.Columns["categories"].HeaderText = "Категория";
                dataGridView1.Columns["description"].HeaderText = "Описание";
                dataGridView1.Columns["stock_quantity"].HeaderText = "Количество на складе";
                dataGridView1.Columns["price"].HeaderText = "Цена";
                dataGridView1.Columns["created_at"].HeaderText = "Дата добавления";
                dataGridView1.Columns["image"].HeaderText = "Изображение";
                dataGridView1.Columns["last_updated"].HeaderText = "Последнее изменение";
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedRow = dataGridView1.SelectedRows[0];
                textBox1.Text = selectedRow.Cells["name"].Value?.ToString() ?? string.Empty;
                txtdescription.Text = selectedRow.Cells["description"].Value?.ToString() ?? string.Empty;
                stock_quantityUpDown1.Value = Convert.ToDecimal(selectedRow.Cells["stock_quantity"].Value);
                maskedprice.Text = selectedRow.Cells["price"].Value?.ToString() ?? string.Empty;
                if (selectedRow.Cells["created_at"].Value != null && selectedRow.Cells["created_at"].Value != DBNull.Value)
                {
                    dateTimePicker1.Value = Convert.ToDateTime(selectedRow.Cells["created_at"].Value);
                }
                if (selectedRow.Cells["image"].Value != null && selectedRow.Cells["image"].Value != DBNull.Value)
                {
                    byte[] imageData = (byte[])selectedRow.Cells["image"].Value;
                    using (var ms = new MemoryStream(imageData))
                    {
                        pictureBox1.Image = Image.FromStream(ms);
                    }
                }
                else
                {
                    pictureBox1.Image = null;
                }
                combocategories.Text = selectedRow.Cells["categories"].Value?.ToString() ?? string.Empty;
                selectedProductId = Convert.ToInt32(selectedRow.Cells["id_products"].Value);
            }
        }

        private byte[] ImageToByteArray(Image imageIn)
        {
            using (var ms = new MemoryStream())
            {
                imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                return ms.ToArray();
            }
        }

        private Image ByteArrayToImage(byte[] byteArray)
        {
            using (var ms = new MemoryStream(byteArray))
            {
                return Image.FromStream(ms);
            }
        }

        private void products_Load(object sender, EventArgs e)
        {
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void button3_Click(object sender, EventArgs e)
        { 
            administrator form2 = new administrator();
            form2.Show();
            this.Hide();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                string name = textBox1.Text;
                string description = txtdescription.Text;
                int stockQuantity = (int)stock_quantityUpDown1.Value;
                decimal price = decimal.Parse(maskedprice.Text);
                DateTime createdAt = dateTimePicker1.Value;
                string categories = combocategories.Text;
                byte[] imageData = null;
                if (pictureBox1.Image != null)
                {
                    using (var ms = new MemoryStream())
                    {
                        pictureBox1.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        imageData = ms.ToArray();
                    }
                }

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("UPDATE public.products SET name = @name, " +
                        "description = @description, stock_quantity = @stockQuantity, price = @price, " +
                        "created_at = @createdAt, image = @image, categories = @categories " +
                        "WHERE id_products = @idProducts", conn))
                    {
                        cmd.Parameters.AddWithValue("name", name);
                        cmd.Parameters.AddWithValue("description", (object)description ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("stockQuantity", stockQuantity);
                        cmd.Parameters.AddWithValue("price", price);
                        cmd.Parameters.AddWithValue("createdAt", createdAt);
                        cmd.Parameters.AddWithValue("image", (object)imageData ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("categories", (object)categories ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("idProducts", selectedProductId);
                        cmd.ExecuteNonQuery();
                    }
                }

                LoadData();
                MessageBox.Show("Данные успешно обновлены.");

                textBox1.Clear();
                txtdescription.Clear();
                combocategories.SelectedIndex = -1;
                maskedprice.Clear();
                stock_quantityUpDown1.Value = stock_quantityUpDown1.Minimum;
                dateTimePicker1.Value = DateTime.Now;
                pictureBox1.Image = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string name = textBox1.Text;
                string description = txtdescription.Text;
                int stockQuantity = (int)stock_quantityUpDown1.Value;
                decimal price = decimal.Parse(maskedprice.Text);
                DateTime createdAt = dateTimePicker1.Value;
                string categories = combocategories.Text;
                byte[] imageData = null;
                if (pictureBox1.Image != null)
                {
                    using (var ms = new MemoryStream())
                    {
                        pictureBox1.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        imageData = ms.ToArray();
                    }
                }

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("INSERT INTO public.products (name, description, stock_quantity, price, created_at, image, categories) " +
                        "VALUES (@name, @description, @stockQuantity, @price, @createdAt, @image, @categories)", conn))
                    {
                        cmd.Parameters.AddWithValue("name", name);
                        cmd.Parameters.AddWithValue("description", (object)description ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("stockQuantity", stockQuantity);
                        cmd.Parameters.AddWithValue("price", price);
                        cmd.Parameters.AddWithValue("createdAt", createdAt);
                        cmd.Parameters.AddWithValue("image", (object)imageData ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("categories", (object)categories ?? DBNull.Value);
                        cmd.ExecuteNonQuery();
                    }
                }

                LoadData(); 
                MessageBox.Show("Продукт успешно добавлен.");

                textBox1.Clear();
                txtdescription.Clear();
                combocategories.SelectedIndex = -1;
                maskedprice.Clear();
                stock_quantityUpDown1.Value = stock_quantityUpDown1.Minimum;
                dateTimePicker1.Value = DateTime.Now;
                pictureBox1.Image = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    pictureBox1.Image = Image.FromFile(openFileDialog.FileName);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string searchText = textBox2.Text.Trim().ToLower();
            if (string.IsNullOrEmpty(searchText))
            {
                LoadData();
                return;
            }

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                DataTable dataTable = (DataTable)dataGridView1.DataSource;

                if (dataTable == null || dataTable.Rows.Count == 0)
                {
                    MessageBox.Show("Нет данных для поиска.");
                    return;
                }

                DataTable filteredTable = dataTable.Clone();

                foreach (DataRow row in dataTable.Rows)
                {
                    bool matchFound = false;
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        if (column.ColumnName == "id_products" || column.DataType == typeof(byte[]))
                            continue;

                        string cellValue = row[column].ToString().ToLower();
                        if (cellValue.Contains(searchText))
                        {
                            matchFound = true;
                            break;
                        }
                    }

                    if (matchFound)
                    {
                        filteredTable.ImportRow(row);
                    }
                }

                dataGridView1.DataSource = filteredTable;

                if (filteredTable.Rows.Count == 0)
                {
                    MessageBox.Show("Товары по вашему запросу не найдены.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске: {ex.Message}");
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void ExportToWord(string filePath)
        {
            try
            {
                var columnsToExport = new Dictionary<string, string>
        {
            { "name", "Наименование" },
            { "description", "Описание" },
            { "stock_quantity", "Количество на складе" },
            { "price", "Цена" },
            { "created_at", "Дата добавления" },
            { "categories", "Категория" }
        };

                using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    SectionProperties sectionProps = new SectionProperties();
                    PageSize pageSize = new PageSize()
                    {
                        Width = 12240,  
                        Height = 15840,
                        Orient = PageOrientationValues.Portrait
                    };

                    PageMargin pageMargin = new PageMargin()
                    {
                        Top = 1000,
                        Right = 700,
                        Bottom = 1000,
                        Left = 700,
                        Header = 500,
                        Footer = 500,
                        Gutter = 0
                    };

                    sectionProps.Append(pageSize);
                    sectionProps.Append(pageMargin);
                    body.Append(sectionProps);

                    Paragraph titleParagraph = new Paragraph(
                        new ParagraphProperties(
                            new Justification() { Val = JustificationValues.Center },
                            new SpacingBetweenLines() { After = "200", Line = "360" }
                        ),
                        new Run(
                            new RunProperties(
                                new Bold(),
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "32" },
                                new RunFonts() { Ascii = "Times New Roman" }
                            ),
                            new Text("Список товаров")
                        )
                    );
                    body.AppendChild(titleParagraph);

                    Table table = new Table();

                    TableProperties tableProperties = new TableProperties(
                        new TableBorders(
                            new TopBorder() { Val = BorderValues.Single, Size = 4 },
                            new BottomBorder() { Val = BorderValues.Single, Size = 4 },
                            new LeftBorder() { Val = BorderValues.Single, Size = 4 },
                            new RightBorder() { Val = BorderValues.Single, Size = 4 },
                            new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 4 },
                            new InsideVerticalBorder() { Val = BorderValues.Single, Size = 4 }
                        ),
                        new TableWidth() { Width = "100%", Type = TableWidthUnitValues.Pct }
                    );
                    table.AppendChild(tableProperties);

                    TableRow headerRow = new TableRow();
                    foreach (var column in columnsToExport)
                    {
                        TableCell cell = new TableCell(
                            new TableCellProperties(
                                new Shading() { Fill = "4472C4" },
                                new VerticalMerge() { Val = MergedCellValues.Restart }
                            ),
                            new Paragraph(
                                new ParagraphProperties(
                                    new Justification() { Val = JustificationValues.Center }
                                ),
                                new Run(
                                    new RunProperties(
                                        new Bold(),
                                        new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "FFFFFF" },
                                        new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "24" },
                                        new RunFonts() { Ascii = "Times New Roman" }
                                    ),
                                    new Text(column.Value)
                                )
                            )
                        );
                        headerRow.AppendChild(cell);
                    }
                    table.AppendChild(headerRow);

                    foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
                    {
                        if (!dgvRow.IsNewRow)
                        {
                            TableRow tableRow = new TableRow();
                            foreach (var column in columnsToExport)
                            {
                                string cellValue = "";

                                if (dgvRow.Cells[column.Key].Value != null)
                                {
                                    cellValue = dgvRow.Cells[column.Key].Value.ToString();
                                }

                                if (column.Key == "created_at" && DateTime.TryParse(cellValue, out DateTime dateValue))
                                {
                                    cellValue = dateValue.ToString("dd.MM.yyyy");
                                }

                                TableCell tableCell = new TableCell(
                                    new TableCellProperties(
                                        new Shading() { Fill = dgvRow.Index % 2 == 0 ? "E9E9E9" : "FFFFFF" }
                                    ),
                                    new Paragraph(
                                        new ParagraphProperties(
                                            new SpacingBetweenLines() { After = "0" }
                                        ),
                                        new Run(
                                            new RunProperties(
                                                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "22" },
                                                new RunFonts() { Ascii = "Times New Roman" }
                                            ),
                                            new Text(cellValue)
                                        )
                                    )
                                );
                                tableRow.AppendChild(tableCell);
                            }
                            table.AppendChild(tableRow);
                        }
                    }
                    body.AppendChild(table);
                    mainPart.Document.Save();
                }

                MessageBox.Show("Данные успешно экспортированы в Word.", "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportToExcel(string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Продукты");

                    var columnsToExport = new Dictionary<string, string>
            {
                { "name", "Наименование" },
                { "description", "Описание" },
                { "stock_quantity", "Количество на складе" },
                { "price", "Цена" },
                { "created_at", "Дата добавления" },
                { "categories", "Категория" }
            };

                    using (var headerRange = worksheet.Cells[1, 1, 1, columnsToExport.Count])
                    {
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(68, 114, 196)); 
                        headerRange.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        headerRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        headerRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        headerRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        headerRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    int colIndex = 1;
                    foreach (var column in columnsToExport)
                    {
                        worksheet.Cells[1, colIndex].Value = column.Value;
                        colIndex++;
                    }

                    int rowIndex = 2;
                    foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
                    {
                        if (!dgvRow.IsNewRow)
                        {
                            colIndex = 1;
                            foreach (var column in columnsToExport)
                            {
                                string cellValue = dgvRow.Cells[column.Key].Value?.ToString() ?? string.Empty;

                                if (column.Key == "created_at" && DateTime.TryParse(cellValue, out DateTime dateValue))
                                {
                                    cellValue = dateValue.ToString("dd.MM.yyyy");
                                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "dd.MM.yyyy";
                                }
                               
                                else if (column.Key == "price" && decimal.TryParse(cellValue, out decimal priceValue))
                                {
                                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "#,##0.00";
                                }

                                worksheet.Cells[rowIndex, colIndex].Value = cellValue;

                                if (rowIndex % 2 == 0)
                                {
                                    worksheet.Cells[rowIndex, colIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    worksheet.Cells[rowIndex, colIndex].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(233, 233, 233)); // Явное указание System.Drawing.Color
                                }
                                colIndex++;
                            }
                            rowIndex++;
                        }
                    }

                    using (var dataRange = worksheet.Cells[1, 1, rowIndex - 1, columnsToExport.Count])
                    {
                        dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.View.FreezePanes(2, 1);
                    package.SaveAs(new FileInfo(filePath));
                    MessageBox.Show("Данные успешно экспортированы в Excel.", "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                List<string> namesToDelete = new List<string>();

                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    if (row.Cells["name"].Value != null)
                    {
                        namesToDelete.Add(row.Cells["name"].Value.ToString());
                    }
                }

                using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    foreach (string nameToDelete in namesToDelete)
                    {
                        string query = "DELETE FROM products WHERE name = @nameToDelete";

                        using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@nameToDelete", nameToDelete);
                            int rowsAffected = command.ExecuteNonQuery();
                        }
                    }
                    LoadData();
                    MessageBox.Show("Данные успешно удалены.");
                    textBox1.Clear();
                    txtdescription.Clear();
                    combocategories.SelectedIndex = -1;
                    maskedprice.Clear();
                    stock_quantityUpDown1.Value = stock_quantityUpDown1.Minimum;
                    dateTimePicker1.Value = DateTime.Now;
                    pictureBox1.Image = null;

                }
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите хотя бы одну строку для удаления.");
            }
        }

        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string name = textBox1.Text;
                string description = txtdescription.Text;
                int stockQuantity = (int)stock_quantityUpDown1.Value;
                decimal price = decimal.Parse(maskedprice.Text);
                DateTime createdAt = dateTimePicker1.Value;
                string categories = combocategories.Text;

                byte[] imageData = null;
                if (pictureBox1.Image != null)
                {
                    using (var ms = new MemoryStream())
                    {
                        pictureBox1.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        imageData = ms.ToArray();
                    }
                }

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("UPDATE public.products SET name = @name, " +
                        "description = @description, stock_quantity = @stockQuantity, price = @price, " +
                        "created_at = @createdAt, image = @image, categories = @categories " +
                        "WHERE id_products = @idProducts", conn))
                    {
                        cmd.Parameters.AddWithValue("name", name);
                        cmd.Parameters.AddWithValue("description", (object)description ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("stockQuantity", stockQuantity);
                        cmd.Parameters.AddWithValue("price", price);
                        cmd.Parameters.AddWithValue("createdAt", createdAt);
                        cmd.Parameters.AddWithValue("image", (object)imageData ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("categories", (object)categories ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("idProducts", selectedProductId);
                        cmd.ExecuteNonQuery();
                    }
                }

                LoadData();
                MessageBox.Show("Данные успешно обновлены.");
                textBox1.Clear();
                txtdescription.Clear();
                combocategories.SelectedIndex = -1;
                maskedprice.Clear();
                stock_quantityUpDown1.Value = stock_quantityUpDown1.Minimum;
                dateTimePicker1.Value = DateTime.Now;
                pictureBox1.Image = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        private void добавитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                string name = textBox1.Text;
                string description = txtdescription.Text;
                int stockQuantity = (int)stock_quantityUpDown1.Value;
                decimal price = decimal.Parse(maskedprice.Text);
                DateTime createdAt = dateTimePicker1.Value;
                string categories = combocategories.Text;
                byte[] imageData = null;
                if (pictureBox1.Image != null)
                {
                    using (var ms = new MemoryStream())
                    {
                        pictureBox1.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        imageData = ms.ToArray();
                    }
                }

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("INSERT INTO public.products (name, description, stock_quantity, price, created_at, image, categories) " +
                        "VALUES (@name, @description, @stockQuantity, @price, @createdAt, @image, @categories)", conn))
                    {
                        cmd.Parameters.AddWithValue("name", name);
                        cmd.Parameters.AddWithValue("description", (object)description ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("stockQuantity", stockQuantity);
                        cmd.Parameters.AddWithValue("price", price);
                        cmd.Parameters.AddWithValue("createdAt", createdAt);
                        cmd.Parameters.AddWithValue("image", (object)imageData ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("categories", (object)categories ?? DBNull.Value);
                        cmd.ExecuteNonQuery();
                    }
                }

                LoadData();
                MessageBox.Show("Продукт успешно добавлен.");

                textBox1.Clear();
                txtdescription.Clear();
                combocategories.SelectedIndex = -1;
                maskedprice.Clear();
                stock_quantityUpDown1.Value = stock_quantityUpDown1.Minimum;
                dateTimePicker1.Value = DateTime.Now;
                pictureBox1.Image = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        private void экспортToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word Documents (*.docx)|*.docx|Excel Files (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "Экспорт данных продуктов";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFilePath = saveFileDialog.FileName;

                    if (selectedFilePath.EndsWith(".docx"))
                    {
                        ExportToWord(selectedFilePath);
                    }
                    else if (selectedFilePath.EndsWith(".xlsx"))
                    {
                        ExportToExcel(selectedFilePath);
                    }
                }
            }
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                textBox1.Clear();
                txtdescription.Clear();
                combocategories.SelectedIndex = -1;
                maskedprice.Clear();
                stock_quantityUpDown1.Value = stock_quantityUpDown1.Minimum;
                dateTimePicker1.Value = DateTime.Now;
                pictureBox1.Image = null;

                LoadData();

                MessageBox.Show("Данные успешно обновлены.", "Обновление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении данных:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            programinfo form2 = new programinfo();
            form2.Show();
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string helpFileName = "Справка по форме товары.pdf";

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

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                button4_Click(sender, e);
                e.Handled = true;
            }
        }
    }
}