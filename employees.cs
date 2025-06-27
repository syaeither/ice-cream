using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using OfficeOpenXml.Style;
using System.IO;
using OfficeOpenXml;
using System.Diagnostics;

namespace ice_cream
{
    public partial class employees : Form
    {
        string connectionString = Program.ConnectionString;
        private int selectedEmployeeId;

        public employees()
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
                var adapter = new NpgsqlDataAdapter("SELECT id_employee, full_name, email, telephone, position FROM employees", conn);
                var table = new DataTable();
                adapter.Fill(table);

                dataGridView1.DataSource = table;

                dataGridView1.Columns["id_employee"].Visible = false;
                dataGridView1.Columns["full_name"].HeaderText = "ФИО";
                dataGridView1.Columns["email"].HeaderText = "Email";
                dataGridView1.Columns["telephone"].HeaderText = "Телефон";
                dataGridView1.Columns["position"].HeaderText = "Должность";
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
                        string query = "DELETE FROM employees WHERE full_name = @nameToDelete";

                        using (NpgsqlCommand command = new NpgsqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@nameToDelete", nameToDelete);
                            int rowsAffected = command.ExecuteNonQuery();
                        }
                    }

                    LoadData();
                    MessageBox.Show("Данные успешно удалены.");
                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox5.Clear();
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
                textBox2.Text = selectedRow.Cells["email"].Value?.ToString() ?? string.Empty;
                textBox3.Text = selectedRow.Cells["position"].Value?.ToString() ?? string.Empty;
                selectedEmployeeId = Convert.ToInt32(selectedRow.Cells["id_employee"].Value);
            }
        }

        private void employees_Load(object sender, EventArgs e)
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
                string fullName = textBox1.Text;
                string email = textBox2.Text;
                string telephone = maskedTextBox1.Text;
                string position = textBox3.Text;

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("UPDATE public.employees SET full_name = @fullName, " +
                        "email = @email, telephone = @telephone, position = @position " +
                        "WHERE id_employee = @idEmployee", conn))
                    {
                        cmd.Parameters.AddWithValue("fullName", fullName);
                        cmd.Parameters.AddWithValue("email", (object)email ?? DBNull.Value); 
                        cmd.Parameters.AddWithValue("telephone", (object)telephone ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("position", (object)position ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("idEmployee", selectedEmployeeId); 
                        cmd.ExecuteNonQuery();
                    }
                }

                LoadData(); 
                MessageBox.Show("Данные успешно обновлены.");

                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
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
                string email = textBox2.Text.Trim(); 
                string telephone = maskedTextBox1.Text.Trim();
                string position = textBox3.Text.Trim();
                string username = textBox4.Text.Trim();
                string password = textBox5.Text.Trim();

                if (string.IsNullOrWhiteSpace(fullName) || string.IsNullOrWhiteSpace(telephone) ||
                    string.IsNullOrWhiteSpace(username)) 
                {
                    MessageBox.Show("Пожалуйста, заполните все обязательные поля.");
                    return;
                }

                using (var conn = new NpgsqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new NpgsqlCommand("INSERT INTO public.employees (full_name, email, telephone, position, username, password) " +
                                                        "VALUES (@fullName, @email, @telephone, @position, @username, @password)", conn))
                    {
                        cmd.Parameters.AddWithValue("fullName", fullName);
                        cmd.Parameters.AddWithValue("email", (object)email ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("telephone", (object)telephone ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("position", (object)position ?? DBNull.Value); 
                        cmd.Parameters.AddWithValue("username", (object)username ?? DBNull.Value); 
                        cmd.Parameters.AddWithValue("password", (object)password ?? DBNull.Value);
                        cmd.ExecuteNonQuery();
                    }
                }

                LoadData();
                MessageBox.Show("Данные успешно добавлены. ");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
                maskedTextBox1.Clear();
            }
            catch (PostgresException ex) when (ex.SqlState == "23505")
            {
                MessageBox.Show("Ошибка: Данный сотрудник уже существует. Пожалуйста, проверьте уникальные поля (например, телефон или username).");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
           
        }

        private void button6_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word Documents (*.docx)|*.docx|Excel Files (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "Экспорт данных сотрудников";

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

        private void ExportToWord(string filePath)
        {
            try
            {
                var columnsToExport = new Dictionary<string, string>
        {
            { "full_name", "ФИО" },
            { "email", "Email" },
            { "telephone", "Телефон" },
            { "position", "Должность" }
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
                            new Text("Список сотрудников")
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
                                string cellValue = dgvRow.Cells[column.Key].Value?.ToString() ?? "";

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

                FileInfo fileInfo = new FileInfo(filePath);

                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Сотрудники");

                    var columnsToExport = new Dictionary<string, string>
            {
                { "full_name", "ФИО" },
                { "email", "Email" },
                { "telephone", "Телефон" },
                { "position", "Должность" }
            };

                    int colIndex = 1;
                    foreach (var column in columnsToExport)
                    {
                        worksheet.Cells[1, colIndex].Value = column.Value;
                        colIndex++;
                    }

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

                    int rowIndex = 2;
                    foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
                    {
                        if (!dgvRow.IsNewRow)
                        {
                            colIndex = 1;
                            foreach (var column in columnsToExport)
                            {
                                worksheet.Cells[rowIndex, colIndex].Value = dgvRow.Cells[column.Key].Value?.ToString() ?? string.Empty;

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

                    if (rowIndex > 2)
                    {
                        using (var dataRange = worksheet.Cells[1, 1, rowIndex - 1, columnsToExport.Count])
                        {
                            dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        }
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    worksheet.View.FreezePanes(2, 1);

                    package.Save();
                    MessageBox.Show("Данные успешно экспортированы в Excel.", "Экспорт завершен",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel:\n{ex.Message}", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            administrator form2 = new administrator();
            form2.Show();
            this.Hide();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();
                textBox5.Clear();
                maskedTextBox1.Clear();
                dataGridView1.ClearSelection();
                LoadData();

                MessageBox.Show("Данные сотрудников успешно обновлены.", "Обновление",MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении данных:\n{ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string searchText = textBox6.Text.Trim().ToLower();

            if (string.IsNullOrEmpty(searchText) || searchText == "поиск по сотрудникам...")
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
                    MessageBox.Show("Нет данных для поиска.", "Поиск",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                DataTable filteredTable = dataTable.Clone();

                foreach (DataRow row in dataTable.Rows)
                {
                    bool matchFound = false;
                    foreach (DataGridViewColumn column in dataGridView1.Columns)
                    {
                        if (!column.Visible || column.Name == "id_employee") continue;

                        string columnName = column.DataPropertyName;
                        string cellValue = row[columnName]?.ToString().ToLower() ?? "";

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
                    MessageBox.Show("Сотрудники по вашему запросу не найдены.", "Поиск", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске:\n{ex.Message}", "Ошибка поиска",MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            if (textBox6.Text.Length > 2 && !textBox6.Text.Contains("Поиск по сотрудникам..."))
            {
                button7_Click(null, null);
            }
            else if (string.IsNullOrWhiteSpace(textBox6.Text) || textBox6.Text == "Поиск по сотрудникам...")
            {
                LoadData();
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            programinfo form2 = new programinfo();
            form2.Show();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            string helpFileName = "Справка по форме сотрудники.pdf";

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