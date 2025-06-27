using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Npgsql;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Diagnostics;

namespace ice_cream
{
    public partial class clients : Form
    {
        string connectionString = Program.ConnectionString;
        private int selectedClientId;

        public clients()
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
                var adapter = new NpgsqlDataAdapter("SELECT * FROM clients", conn);
                var table = new DataTable();
                adapter.Fill(table);

                dataGridView1.DataSource = table;

                dataGridView1.Columns["id_client"].Visible = false;
                dataGridView1.Columns["full_name"].HeaderText = "ФИО";
                dataGridView1.Columns["telephone"].HeaderText = "Телефон";
                dataGridView1.Columns["created_at"].HeaderText = "Дата регистрации";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                List<string> namesToDelete = new List<string>();

                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    if (row.Cells["full_name"].Value != null)
                    {
                        namesToDelete.Add(row.Cells["full_name"].Value.ToString());
                    }
                }

                using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
                {
                    connection.Open();
                    foreach (string nameToDelete in namesToDelete)
                    {
                        string query = "DELETE FROM clients WHERE full_name = @nameToDelete";

                        using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@nameToDelete", nameToDelete);
                            int rowsAffected = command.ExecuteNonQuery();
                        }
                    }
                    LoadData();
                    MessageBox.Show("Данные успешно удалены.");
                    textBox1.Clear();
                    maskedTextBox1.Clear();

                }
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите хотя бы одну строку для удаления.");
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedRow = dataGridView1.SelectedRows[0];

                textBox1.Text = selectedRow.Cells["full_name"].Value.ToString();
                maskedTextBox1.Text = selectedRow.Cells["telephone"].Value?.ToString() ?? string.Empty;

                dateTimePicker1.Value = selectedRow.Cells["created_at"].Value != DBNull.Value
                    ? Convert.ToDateTime(selectedRow.Cells["created_at"].Value)
                    : DateTime.Now;

                selectedClientId = Convert.ToInt32(selectedRow.Cells["id_client"].Value);
            }
        }

        private void clients_Load(object sender, EventArgs e)
        {
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }    

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {             
                string fullName = textBox1.Text;
                string telephone = maskedTextBox1.Text;
                DateTime? createdAt = DateTime.Now;
              
                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("UPDATE public.clients SET full_name = @fullName, " +
                        "telephone = @telephone, created_at = @createdAt WHERE id_client = @idClient", conn))
                    {
                        cmd.Parameters.AddWithValue("fullName", fullName);
                        cmd.Parameters.AddWithValue("telephone", (object)telephone ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("createdAt", createdAt); 
                        cmd.Parameters.AddWithValue("idClient", selectedClientId);
                        cmd.ExecuteNonQuery();
                    }
                }

                LoadData(); 
                MessageBox.Show("Данные успешно обновлены.");
               
                textBox1.Clear();
                maskedTextBox1.Clear();
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
                string fullName = textBox1.Text.Trim();
                string telephone = maskedTextBox1.Text.Trim();
                DateTime createdAt = DateTime.Now; 
                
                if (string.IsNullOrWhiteSpace(fullName) || string.IsNullOrWhiteSpace(telephone))
                {
                    MessageBox.Show("Пожалуйста, заполните все обязательные поля.");
                    return;
                }

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("INSERT INTO public.clients (full_name, telephone, created_at) " +
                                                        "VALUES (@fullName, @telephone, @createdAt)", conn))
                    {
                        cmd.Parameters.AddWithValue("fullName", fullName);
                        cmd.Parameters.AddWithValue("telephone", telephone);
                        cmd.Parameters.AddWithValue("createdAt", createdAt);
                        cmd.ExecuteNonQuery();
                    }
                }
                LoadData();
                MessageBox.Show("Данные успешно добавлены. ");
                textBox1.Clear();
                maskedTextBox1.Clear();
            }
            catch (PostgresException ex) when (ex.SqlState == "23505")
            {
                MessageBox.Show("Ошибка: Данный клиент уже существует. Пожалуйста, проверьте уникальные поля (например, телефон).");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word Documents (*.docx)|*.docx|Excel Files (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "Экспорт данных клиентов";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFilePath = saveFileDialog.FileName;

                    if (selectedFilePath.EndsWith(".docx"))
                    {
                        ExportClientsToWord(selectedFilePath);
                    }
                    else if (selectedFilePath.EndsWith(".xlsx"))
                    {
                        ExportClientsToExcel(selectedFilePath);
                    }
                }
            }
        }

        private void ExportClientsToWord(string filePath)
        {
            try
            {
                var columnsToExport = new Dictionary<string, string>
        {
            { "full_name", "ФИО" },
            { "telephone", "Телефон" },
            { "created_at", "Дата регистрации" }
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
                            new Text("Список клиентов")
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

        private void ExportClientsToExcel(string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Клиенты");

                    var columnsToExport = new Dictionary<string, string>
            {
                { "full_name", "ФИО" },
                { "telephone", "Телефон" },
                { "created_at", "Дата регистрации" }
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
                                    worksheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = "dd.MM.yyyy";
                                    worksheet.Cells[rowIndex, colIndex].Value = dateValue;
                                }
                                else
                                {
                                    worksheet.Cells[rowIndex, colIndex].Value = cellValue;
                                }

                                if (rowIndex % 2 == 0)
                                {
                                    worksheet.Cells[rowIndex, colIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    worksheet.Cells[rowIndex, colIndex].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(233, 233, 233));
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

        private void button8_Click(object sender, EventArgs e)
        {
            administrator form2 = new administrator();
            form2.Show();
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Кнопка "Обновить"
            try
            {
                textBox1.Clear();
                maskedTextBox1.Clear();
                LoadData();

                MessageBox.Show("Данные успешно обновлены.", "Обновление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении данных:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            // Кнопка "Поиск"
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
                        if (column.ColumnName == "id_client")
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
                    MessageBox.Show("Клиенты по вашему запросу не найдены.");
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

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                button7_Click(sender, e);
                e.Handled = true;
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            programinfo form2 = new programinfo();
            form2.Show();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string helpFileName = "Справка по формме клиенты.pdf";

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