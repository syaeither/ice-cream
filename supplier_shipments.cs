using Npgsql;
using System;
using System.Data;
using System.Windows.Forms;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using JustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;
using TableWidthUnitValues = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues;
using System.Diagnostics;

namespace ice_cream
{
    public partial class supplier_shipments : Form
    {
        string connectionString = Program.ConnectionString;
        private int selectedShipmentId;

        public supplier_shipments()
        {
            InitializeComponent();
            LoadData();
            dataGridView1.AllowUserToAddRows = false;
        }

        private void LoadData()
        {
            using (var conn = new NpgsqlConnection(connectionString))
            {
                conn.Open();
                var adapter = new NpgsqlDataAdapter("SELECT ss.id_supplier, ss.shipment_date, ss.quantity, ss.fullname_supplier, " +
                                                    "ss.telephone, ss.email, ss.company_name, p.name AS product_name " +
                                                    "FROM supplier ss " +
                                                    "JOIN products p ON ss.code_products = p.id_products", conn);
                var table = new DataTable();
                adapter.Fill(table);

                dataGridView1.DataSource = table;
                dataGridView1.Columns["id_supplier"].Visible = false;
                dataGridView1.Columns["fullname_supplier"].HeaderText = "ФИО";
                dataGridView1.Columns["company_name"].HeaderText = "Наименование компании";
                dataGridView1.Columns["telephone"].HeaderText = "Телефон";
                dataGridView1.Columns["email"].HeaderText = "Email";
                dataGridView1.Columns["product_name"].HeaderText = "Наименование продукта";
                dataGridView1.Columns["shipment_date"].HeaderText = "Дата отправки";
                dataGridView1.Columns["quantity"].HeaderText = "Количество";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedRow = dataGridView1.SelectedRows[0];
                int selectedId = Convert.ToInt32(selectedRow.Cells["id_supplier"].Value);

                using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "DELETE FROM supplier WHERE id_supplier = @selectedId";
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@selectedId", selectedId);
                        command.ExecuteNonQuery();
                    }
                }

                LoadData();
                ClearFields();
                MessageBox.Show("Данные успешно удалены.");
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите хотя бы одну строку для удаления.");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(textBoxFullName.Text) ||
                    string.IsNullOrWhiteSpace(textBoxCompanyName.Text) ||
                    string.IsNullOrWhiteSpace(maskedTextBoxTelephone.Text) ||
                    comboBoxProductName.SelectedIndex == -1 ||
                    string.IsNullOrWhiteSpace(textBoxQuantity.Text))
                {
                    MessageBox.Show("Пожалуйста, заполните все обязательные поля: ФИО, наименование компании, телефон, продукт и количество.");
                    return;
                }

                if (!int.TryParse(textBoxQuantity.Text, out int quantity) || quantity <= 0)
                {
                    MessageBox.Show("Пожалуйста, введите корректное количество (целое число больше 0).");
                    return;
                }

                string fullName = textBoxFullName.Text.Trim();
                string companyName = textBoxCompanyName.Text.Trim();
                string telephone = maskedTextBoxTelephone.Text.Trim();
                string email = textBoxEmail.Text.Trim();
                string productName = comboBoxProductName.SelectedItem.ToString();

                int productId;
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("SELECT id_products FROM products WHERE name = @productName", conn))
                    {
                        cmd.Parameters.AddWithValue("productName", productName);
                        productId = (int)cmd.ExecuteScalar();
                    }
                }

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand(
                        "INSERT INTO public.supplier (fullname_supplier, company_name, telephone, email, code_products, quantity, shipment_date) " +
                        "VALUES (@fullName, @companyName, @telephone, @email, @productId, @quantity, @shipmentDate)", conn))
                    {
                        cmd.Parameters.AddWithValue("fullName", fullName);
                        cmd.Parameters.AddWithValue("companyName", companyName);
                        cmd.Parameters.AddWithValue("telephone", telephone);
                        cmd.Parameters.AddWithValue("email", email);
                        cmd.Parameters.AddWithValue("productId", productId);
                        cmd.Parameters.AddWithValue("quantity", quantity);
                        cmd.Parameters.AddWithValue("shipmentDate", DateTime.Now);
                        cmd.ExecuteNonQuery();
                    }
                }

                LoadData();
                ClearFields();
                MessageBox.Show("Данные успешно добавлены.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                string fullName = textBoxFullName.Text.Trim();
                string companyName = textBoxCompanyName.Text.Trim();
                string telephone = maskedTextBoxTelephone.Text.Trim();
                string email = textBoxEmail.Text.Trim();
                string productName = comboBoxProductName.SelectedItem.ToString();
                int quantity = int.Parse(textBoxQuantity.Text.Trim());
                DateTime shipmentDate = dateTimePickerShipmentDate.Value;

                int productId;
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("SELECT id_products FROM products WHERE name = @productName", conn))
                    {
                        cmd.Parameters.AddWithValue("productName", productName);
                        productId = (int)cmd.ExecuteScalar();
                    }
                }

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("UPDATE public.supplier SET fullname_supplier = @fullName, " +
                        "company_name = @companyName, telephone = @telephone, email = @email, " +
                        "code_products = @productId, quantity = @quantity, shipment_date = @shipmentDate " +
                        "WHERE id_supplier = @id", conn))
                    {
                        cmd.Parameters.AddWithValue("fullName", fullName);
                        cmd.Parameters.AddWithValue("companyName", companyName);
                        cmd.Parameters.AddWithValue("telephone", telephone);
                        cmd.Parameters.AddWithValue("email", email);
                        cmd.Parameters.AddWithValue("productId", productId);
                        cmd.Parameters.AddWithValue("quantity", quantity);
                        cmd.Parameters.AddWithValue("shipmentDate", shipmentDate);
                        cmd.Parameters.AddWithValue("id", selectedShipmentId);
                        cmd.ExecuteNonQuery();
                    }
                }

                LoadData();
                ClearFields();
                MessageBox.Show("Данные успешно обновлены.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
            {
                connection.Open();
                string supplierName = textBoxFullName.Text.Trim();
                string query;

                if (string.IsNullOrWhiteSpace(supplierName))
                {
                    query = "SELECT id_supplier AS \"ID\", fullname_supplier AS \"ФИО\", company_name AS \"Компания\", " +
                            "telephone AS \"Телефон\", shipment_date AS \"Дата отгрузки\", quantity AS \"Количество\" " +
                            "FROM supplier";
                }
                else
                {
                    query = "SELECT id_supplier AS \"ID\", fullname_supplier AS \"ФИО\", company_name AS \"Компания\", " +
                            "telephone AS \"Телефон\", shipment_date AS \"Дата отгрузки\", quantity AS \"Количество\" " +
                            "FROM supplier WHERE fullname_supplier ILIKE @fullname_supplier";
                }

                using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                {
                    if (!string.IsNullOrWhiteSpace(supplierName))
                    {
                        command.Parameters.AddWithValue("fullname_supplier", "%" + supplierName + "%");
                    }

                    NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(command);
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    if (dataTable.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = dataTable;
                    }
                    else
                    {
                        MessageBox.Show("Ничего не найдено. Пожалуйста, попробуйте другой запрос.", "Ошибка поиска", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        dataGridView1.DataSource = null;
                    }
                }
                connection.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            administrator form2 = new administrator();
            form2.Show();
            this.Hide();
        }

        private void ClearFields()
        {
            textBoxFullName.Clear();
            textBoxCompanyName.Clear();
            maskedTextBoxTelephone.Clear();
            textBoxEmail.Clear();
            textBoxQuantity.Clear();
            comboBoxProductName.SelectedIndex = -1;
            dateTimePickerShipmentDate.Value = DateTime.Now;
        }

        private void supplier_shipments_Load(object sender, EventArgs e)
        {
            using (var conn = new NpgsqlConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    var cmd = new NpgsqlCommand("SELECT name FROM products", conn);
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            comboBoxProductName.Items.Add(reader.GetString(0));
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка при загрузке продуктов: " + ex.Message);
                }
            }

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedRow = dataGridView1.SelectedRows[0];

                textBoxFullName.Text = selectedRow.Cells["fullname_supplier"].Value.ToString();
                textBoxCompanyName.Text = selectedRow.Cells["company_name"].Value?.ToString() ?? string.Empty;
                maskedTextBoxTelephone.Text = selectedRow.Cells["telephone"].Value?.ToString() ?? string.Empty;
                textBoxEmail.Text = selectedRow.Cells["email"].Value?.ToString() ?? string.Empty;
                textBoxQuantity.Text = selectedRow.Cells["quantity"].Value?.ToString() ?? string.Empty;
                dateTimePickerShipmentDate.Value = Convert.ToDateTime(selectedRow.Cells["shipment_date"].Value);
                string productName = selectedRow.Cells["product_name"].Value.ToString();
                comboBoxProductName.SelectedItem = productName;
                selectedShipmentId = Convert.ToInt32(selectedRow.Cells["id_supplier"].Value);
            }
        }

        private void ExportToExcel(string filePath)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Поставщики");

                    worksheet.Cell(1, 1).Value = "ФИО поставщика";
                    worksheet.Cell(1, 2).Value = "Наименование компании";
                    worksheet.Cell(1, 3).Value = "Телефон";
                    worksheet.Cell(1, 4).Value = "Email";
                    worksheet.Cell(1, 5).Value = "Наименование продукта";
                    worksheet.Cell(1, 6).Value = "Дата отправки";
                    worksheet.Cell(1, 7).Value = "Количество";

                    var headerRange = worksheet.Range("A1:G1");
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(68, 114, 196);
                    headerRange.Style.Font.FontColor = XLColor.White;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if (!dataGridView1.Rows[i].IsNewRow)
                        {
                            worksheet.Cell(i + 2, 1).Value = dataGridView1.Rows[i].Cells["fullname_supplier"].Value?.ToString();
                            worksheet.Cell(i + 2, 2).Value = dataGridView1.Rows[i].Cells["company_name"].Value?.ToString();
                            worksheet.Cell(i + 2, 3).Value = dataGridView1.Rows[i].Cells["telephone"].Value?.ToString();
                            worksheet.Cell(i + 2, 4).Value = dataGridView1.Rows[i].Cells["email"].Value?.ToString();
                            worksheet.Cell(i + 2, 5).Value = dataGridView1.Rows[i].Cells["product_name"].Value?.ToString();

                            if (dataGridView1.Rows[i].Cells["shipment_date"].Value != null &&
                                DateTime.TryParse(dataGridView1.Rows[i].Cells["shipment_date"].Value.ToString(), out DateTime date))
                            {
                                worksheet.Cell(i + 2, 6).Value = date;
                                worksheet.Cell(i + 2, 6).Style.DateFormat.Format = "dd.MM.yyyy";
                            }

                            worksheet.Cell(i + 2, 7).Value = dataGridView1.Rows[i].Cells["quantity"].Value?.ToString();

                            var rowRange = worksheet.Range($"A{i + 2}:G{i + 2}");
                            rowRange.Style.Fill.BackgroundColor = i % 2 == 0
                                ? XLColor.FromArgb(233, 233, 233)
                                : XLColor.White;
                            rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            rowRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        }
                    }
                    worksheet.Columns().AdjustToContents(5);
                    worksheet.SheetView.Freeze(1, 0);
                    workbook.SaveAs(filePath);
                }
                MessageBox.Show("Данные успешно экспортированы в Excel.", "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportToWord(string filePath)
        {
            try
            {
                var columnsToExport = new Dictionary<string, string>
                {
                    { "fullname_supplier", "ФИО поставщика" },
                    { "company_name", "Наименование компании" },
                    { "telephone", "Телефон" },
                    { "email", "Email" },
                    { "product_name", "Наименование продукта" },
                    { "shipment_date", "Дата отправки" },
                    { "quantity", "Количество" }
                };

                using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
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

                    DocumentFormat.OpenXml.Wordprocessing.Paragraph titleParagraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new ParagraphProperties(
                            new DocumentFormat.OpenXml.Wordprocessing.Justification() { Val = JustificationValues.Center },
                            new SpacingBetweenLines() { After = "200", Line = "360" }
                        ),
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new RunProperties(
                                new Bold(),
                                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "32" },
                                new RunFonts() { Ascii = "Times New Roman" }
                            ),
                            new Text("Данные о поставках")
                        )
                    );
                    body.AppendChild(titleParagraph);

                    DocumentFormat.OpenXml.Wordprocessing.Table table = new DocumentFormat.OpenXml.Wordprocessing.Table();

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
                            new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                                new ParagraphProperties(
                                    new DocumentFormat.OpenXml.Wordprocessing.Justification() { Val = JustificationValues.Center }
                                ),
                                new DocumentFormat.OpenXml.Wordprocessing.Run(
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

                                if (column.Key == "shipment_date" && DateTime.TryParse(cellValue, out DateTime dateValue))
                                {
                                    cellValue = dateValue.ToString("dd.MM.yyyy");
                                }

                                TableCell tableCell = new TableCell(
                                    new TableCellProperties(
                                        new Shading() { Fill = dgvRow.Index % 2 == 0 ? "E9E9E9" : "FFFFFF" }
                                    ),
                                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                                        new ParagraphProperties(
                                            new SpacingBetweenLines() { After = "0" }
                                        ),
                                        new DocumentFormat.OpenXml.Wordprocessing.Run(
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

                MessageBox.Show("Данные успешно экспортированы в Word.", "Экспорт завершен",
                               MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word:\n{ex.Message}", "Ошибка",
                               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
                administrator form2 = new administrator();
                form2.Show();
                this.Hide();
        }

        private void добавитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(textBoxFullName.Text) ||
                    string.IsNullOrWhiteSpace(textBoxCompanyName.Text) ||
                    string.IsNullOrWhiteSpace(maskedTextBoxTelephone.Text) ||
                    comboBoxProductName.SelectedIndex == -1 ||
                    string.IsNullOrWhiteSpace(textBoxQuantity.Text))
                {
                    MessageBox.Show("Пожалуйста, заполните все обязательные поля: ФИО, наименование компании, телефон, продукт и количество.");
                    return;
                }

                if (!int.TryParse(textBoxQuantity.Text, out int quantity) || quantity <= 0)
                {
                    MessageBox.Show("Пожалуйста, введите корректное количество (целое число больше 0).");
                    return;
                }

                string fullName = textBoxFullName.Text.Trim();
                string companyName = textBoxCompanyName.Text.Trim();
                string telephone = maskedTextBoxTelephone.Text.Trim();
                string email = textBoxEmail.Text.Trim();
                string productName = comboBoxProductName.SelectedItem.ToString();

                int productId;
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("SELECT id_products FROM products WHERE name = @productName", conn))
                    {
                        cmd.Parameters.AddWithValue("productName", productName);
                        productId = (int)cmd.ExecuteScalar();
                    }
                }

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand(
                        "INSERT INTO public.supplier (fullname_supplier, company_name, telephone, email, code_products, quantity, shipment_date) " +
                        "VALUES (@fullName, @companyName, @telephone, @email, @productId, @quantity, @shipmentDate)", conn))
                    {
                        cmd.Parameters.AddWithValue("fullName", fullName);
                        cmd.Parameters.AddWithValue("companyName", companyName);
                        cmd.Parameters.AddWithValue("telephone", telephone);
                        cmd.Parameters.AddWithValue("email", email);
                        cmd.Parameters.AddWithValue("productId", productId);
                        cmd.Parameters.AddWithValue("quantity", quantity);
                        cmd.Parameters.AddWithValue("shipmentDate", DateTime.Now);
                        cmd.ExecuteNonQuery();
                    }
                }
                LoadData();
                ClearFields();
                MessageBox.Show("Данные успешно добавлены.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string fullName = textBoxFullName.Text.Trim();
                string companyName = textBoxCompanyName.Text.Trim();
                string telephone = maskedTextBoxTelephone.Text.Trim();
                string email = textBoxEmail.Text.Trim();
                string productName = comboBoxProductName.SelectedItem.ToString();
                int quantity = int.Parse(textBoxQuantity.Text.Trim());
                DateTime shipmentDate = dateTimePickerShipmentDate.Value;

                int productId;
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("SELECT id_products FROM products WHERE name = @productName", conn))
                    {
                        cmd.Parameters.AddWithValue("productName", productName);
                        productId = (int)cmd.ExecuteScalar();
                    }
                }

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("UPDATE public.supplier SET fullname_supplier = @fullName, " +
                        "company_name = @companyName, telephone = @telephone, email = @email, " +
                        "code_products = @productId, quantity = @quantity, shipment_date = @shipmentDate " +
                        "WHERE id_supplier = @id", conn))
                    {
                        cmd.Parameters.AddWithValue("fullName", fullName);
                        cmd.Parameters.AddWithValue("companyName", companyName);
                        cmd.Parameters.AddWithValue("telephone", telephone);
                        cmd.Parameters.AddWithValue("email", email);
                        cmd.Parameters.AddWithValue("productId", productId);
                        cmd.Parameters.AddWithValue("quantity", quantity);
                        cmd.Parameters.AddWithValue("shipmentDate", shipmentDate);
                        cmd.Parameters.AddWithValue("id", selectedShipmentId);
                        cmd.ExecuteNonQuery();
                    }
                }

                LoadData();
                ClearFields();
                MessageBox.Show("Данные успешно обновлены.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedRow = dataGridView1.SelectedRows[0];
                int selectedId = Convert.ToInt32(selectedRow.Cells["id_supplier"].Value);

                using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "DELETE FROM supplier WHERE id_supplier = @selectedId";
                    using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@selectedId", selectedId);
                        command.ExecuteNonQuery();
                    }
                }

                LoadData();
                MessageBox.Show("Данные успешно удалены.");
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите хотя бы одну строку для удаления.");
            }
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            programinfo form2 = new programinfo();
            form2.Show();
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                ClearFields();
                LoadData();

                MessageBox.Show("Данные успешно обновлены.", "Обновление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении данных:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void экспортToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word Documents (*.docx)|*.docx|Excel Files (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "Экспорт данных поставщиков";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog.FileName;

                    if (filePath.EndsWith(".docx"))
                    {
                        ExportToWord(filePath);
                    }
                    else if (filePath.EndsWith(".xlsx"))
                    {
                        ExportToExcel(filePath);
                    }
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string searchText = textBox2.Text.Trim();

            if (string.IsNullOrEmpty(searchText))
            {
                LoadData();
                return;
            }

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();

                    string query = @"SELECT ss.id_supplier, ss.shipment_date, ss.quantity, 
                           ss.fullname_supplier, ss.company_name, ss.telephone, ss.email, p.name AS product_name
                           FROM supplier ss
                           JOIN products p ON ss.code_products = p.id_products
                           WHERE ss.fullname_supplier ILIKE @search OR
                                 ss.company_name ILIKE @search OR
                                 ss.telephone ILIKE @search OR
                                 ss.email ILIKE @search OR
                                 p.name ILIKE @search OR
                                 CAST(ss.quantity AS TEXT) ILIKE @search";

                    var adapter = new NpgsqlDataAdapter(query, conn);
                    adapter.SelectCommand.Parameters.AddWithValue("@search", $"%{searchText}%");

                    var table = new DataTable();
                    adapter.Fill(table);

                    if (table.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = table;
                        dataGridView1.Columns["id_supplier"].Visible = false;
                        dataGridView1.Columns["fullname_supplier"].HeaderText = "ФИО";
                        dataGridView1.Columns["company_name"].HeaderText = "Наименование компании";
                        dataGridView1.Columns["telephone"].HeaderText = "Телефон";
                        dataGridView1.Columns["email"].HeaderText = "Email";
                        dataGridView1.Columns["product_name"].HeaderText = "Наименование продукта";
                        dataGridView1.Columns["shipment_date"].HeaderText = "Дата отправки";
                        dataGridView1.Columns["quantity"].HeaderText = "Количество";
                    }
                    else
                    {
                        MessageBox.Show("Поставки по вашему запросу не найдены.");
                        LoadData();
                    }
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

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string helpFileName = "Справка по форме поставок.pdf";

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