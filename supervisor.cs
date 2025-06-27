using System;
using System.Data;
using System.Windows.Forms;
using Npgsql;
using ClosedXML.Excel;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;

namespace ice_cream
{
    public partial class supervisor : Form
    {
        private NpgsqlConnection connection;
        private DataTable reportData;
        private System.Threading.Timer _searchTimer;
        private string _lastSearchText = "";

        public supervisor(string fullName = "")
        {
            InitializeComponent();
            dataGridView1.AllowUserToAddRows = false;
            string connectionString = Program.ConnectionString;
            connection = new NpgsqlConnection(connectionString);
            try
            {
                connection.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка подключения к базе данных: " + ex.Message);
            }

            Load += (sender, e) => {
                string appPath = Assembly.GetExecutingAssembly().Location;
                label6.Text = $"{Path.GetDirectoryName(appPath)}";

                label6.Text = string.IsNullOrEmpty(UserSession.FullName) ?
                    "Не авторизован" :
                    $"{UserSession.FullName}";
            };

            textBox2.TextChanged += textBox2_TextChanged;
            textBox2.KeyDown += textBox2_KeyDown;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Authorization form2 = new Authorization();
            form2.Show();
            this.Hide();
        }

        private void supervisor_Load(object sender, EventArgs e)
        {
            string[] reportTypes = { "Поставки", "Сотрудники", "Продажи" };
            comboBox1.DataSource = reportTypes;

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime startDate = dateTimePicker1.Value;
            DateTime endDate = dateTimePicker2.Value;
            string reportType = comboBox1.Text;

            if (reportType == "Поставки")
            {
                reportData = GetShipmentData(startDate, endDate);
                decimal totalSum = CalculateTotalSum(reportData, "Стоимость поставок");
                textBox1.Text = totalSum.ToString("C");
            }
            else if (reportType == "Сотрудники")
            {
                reportData = GetEmployeePerformanceReport(startDate, endDate);

                // Рассчитываем общую сумму для отчета по сотрудникам
                decimal totalSum = 0;
                foreach (DataRow row in reportData.Rows)
                {
                    // Для администраторов суммируем общую сумму заказов
                    if (row["Должность"].ToString() == "Администратор")
                    {
                        totalSum += Convert.ToDecimal(row["Общая сумма заказов"]);
                    }
                    // Для продавцов суммируем общую сумму продаж
                    else if (row["Должность"].ToString() == "Продавец")
                    {
                        totalSum += Convert.ToDecimal(row["Общая сумма продаж"]);
                    }
                }
                textBox1.Text = totalSum.ToString("C");
            }
            else if (reportType == "Продажи")
            {
                reportData = GetSalesData(startDate, endDate);
                decimal totalSum = CalculateTotalSum(reportData, "Общая сумма");
                textBox1.Text = totalSum.ToString("C");
            }

            dataGridView1.DataSource = reportData;
        }

        private DataTable GetShipmentData(DateTime startDate, DateTime endDate)
        {
            DataTable dt = new DataTable();
            string query = @"
            SELECT 
                ss.shipment_date AS ""Дата поставки"",
                ss.quantity AS ""Количество"",
                p.name AS ""Наименование товара"",
                p.price AS ""Цена товара"",
                ss.fullname_supplier AS ""ФИО поставщика"",
                ss.telephone AS ""Телефон"",
                (ss.quantity * p.price) AS ""Стоимость поставок""  -- Новый столбец для стоимости поставок
            FROM 
                public.supplier ss
            JOIN 
                public.products p ON ss.code_products = p.id_products
            WHERE 
                ss.shipment_date BETWEEN @startDate AND @endDate";

            using (NpgsqlCommand cmd = new NpgsqlCommand(query, connection))
            {
                cmd.Parameters.AddWithValue("startDate", startDate);
                cmd.Parameters.AddWithValue("endDate", endDate);
                using (NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(cmd))
                {
                    adapter.Fill(dt);
                }
            }
            return dt;
        }

        private DataTable GetSalesData(DateTime startDate, DateTime endDate)
        {
            DataTable dt = new DataTable();
            string query = @"
            SELECT 
                s.order_date AS ""Дата продажи"",
                e.full_name AS ""ФИО сотрудника"",
                SUM(s.total_amount) AS ""Общая сумма"",
                COUNT(s.id_sale) AS ""Количество продаж""
            FROM 
                public.sales s
            JOIN 
                public.employees e ON s.code_employee = e.id_employee
            WHERE 
                s.order_date BETWEEN @startDate AND @endDate
            GROUP BY 
                s.order_date, e.full_name";

            using (NpgsqlCommand cmd = new NpgsqlCommand(query, connection))
            {
                cmd.Parameters.AddWithValue("startDate", startDate);
                cmd.Parameters.AddWithValue("endDate", endDate);
                using (NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(cmd))
                {
                    adapter.Fill(dt);
                }
            }
            return dt;
        }

        private DataTable GetEmployeePerformanceReport(DateTime startDate, DateTime endDate)
        {
            DataTable reportTable = new DataTable();

            string query = @"
    SELECT 
        e.full_name AS ""ФИО сотрудника"",
        e.position AS ""Должность"",
        
        -- Для администраторов: только данные по заказам
        CASE WHEN e.position = 'Администратор' 
             THEN COUNT(DISTINCT o.id_order) ELSE 0 END AS ""Всего заказов"",
        CASE WHEN e.position = 'Администратор' 
             THEN COUNT(DISTINCT CASE WHEN o.status = 'В процессе' THEN o.id_order END) ELSE 0 END AS ""Заказов в процессе"",
        CASE WHEN e.position = 'Администратор' 
             THEN COUNT(DISTINCT CASE WHEN o.status = 'Завершен' THEN o.id_order END) ELSE 0 END AS ""Завершенных заказов"",
        CASE WHEN e.position = 'Администратор' 
             THEN COALESCE(SUM(o.total_cost), 0) ELSE 0 END AS ""Общая сумма заказов"",
        CASE WHEN e.position = 'Администратор' 
             THEN COALESCE(SUM(CASE WHEN o.status = 'В процессе' THEN o.total_cost END), 0) ELSE 0 END AS ""Сумма заказов в процессе"",
        CASE WHEN e.position = 'Администратор' 
             THEN COALESCE(SUM(CASE WHEN o.status = 'Завершен' THEN o.total_cost END), 0) ELSE 0 END AS ""Сумма завершенных заказов"",
        
        -- Для продавцов: только данные по продажам
        CASE WHEN e.position = 'Продавец' 
             THEN COALESCE(SUM(s.total_amount), 0) ELSE 0 END AS ""Общая сумма продаж"",
        
        -- Для руководителей: пустые значения (они не работают ни с заказами, ни с продажами)
        0 AS ""Неприменимо для руководителей""
    FROM 
        public.employees e
    LEFT JOIN 
        public.orders o ON e.id_employee = o.code_employee
        AND o.date_orders BETWEEN @startDate AND @endDate
        AND e.position = 'Администратор' -- Только для администраторов
    LEFT JOIN 
        public.sales s ON e.id_employee = s.code_employee
        AND s.order_date BETWEEN @startDate AND @endDate
        AND e.position = 'Продавец' -- Только для продавцов
    WHERE
        e.position IN ('Администратор', 'Продавец', 'Руководитель')
    GROUP BY 
        e.full_name, e.position
    ORDER BY 
        e.position, e.full_name";

            using (NpgsqlCommand cmd = new NpgsqlCommand(query, connection))
            {
                cmd.Parameters.AddWithValue("startDate", startDate);
                cmd.Parameters.AddWithValue("endDate", endDate);

                using (NpgsqlDataAdapter adapter = new NpgsqlDataAdapter(cmd))
                {
                    adapter.Fill(reportTable);
                }
            }

            // Удаляем техническую колонку для руководителей
            reportTable.Columns.Remove("Неприменимо для руководителей");

            return reportTable;
        }

        private decimal CalculateTotalSum(DataTable reportData, string columnName)
        {
            decimal totalSum = 0;

            if (reportData.Columns.Contains(columnName))
            {
                foreach (DataRow row in reportData.Rows)
                {
                    if (row[columnName] != DBNull.Value)
                    {
                        totalSum += Convert.ToDecimal(row[columnName]);
                    }
                }
            }

            return totalSum;
        }

        private void ExportToWord(string filePath)
        {
            try
            {
                DateTime startDate = dateTimePicker1.Value;
                DateTime endDate = dateTimePicker2.Value;
                string reportTitle;
                switch (comboBox1.Text)
                {
                    case "Поставки":
                        reportTitle = "Отчет по поставкам";
                        break;
                    case "Сотрудники":
                        reportTitle = "Отчет по сотрудникам";
                        break;
                    case "Продажи":
                        reportTitle = "Отчет по продажам";
                        break;
                    default:
                        reportTitle = "Отчет";
                        break;
                }

                using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    SectionProperties sectionProps = new SectionProperties();
                    PageSize pageSize = new PageSize()
                    {
                        Width = 15840,
                        Height = 12240,
                        Orient = PageOrientationValues.Landscape
                    };

                    PageMargin pageMargin = new PageMargin()
                    {
                        Top = 500,
                        Right = 500,
                        Bottom = 500,
                        Left = 500,
                        Header = 0,
                        Footer = 0,
                        Gutter = 0
                    };

                    sectionProps.Append(pageSize);
                    sectionProps.Append(pageMargin);
                    body.Append(sectionProps);

                    // Main title
                    Paragraph titleParagraph = new Paragraph(
                        new ParagraphProperties(
                            new SpacingBetweenLines()
                            {
                                After = "0",
                                Line = "360",
                                LineRule = LineSpacingRuleValues.Auto
                            }
                        ),
                        new Run(
                            new RunProperties(
                                new Bold(),
                                new FontSize() { Val = "32" }
                            ),
                            new Text(reportTitle)
                        )
                    );
                    body.AppendChild(titleParagraph);

                    // Date period
                    Paragraph dateParagraph = new Paragraph(
                        new ParagraphProperties(
                            new SpacingBetweenLines()
                            {
                                After = "0",
                                Line = "240",
                                LineRule = LineSpacingRuleValues.Auto
                            }
                        ),
                        new Run(
                            new RunProperties(
                                new FontSize() { Val = "24" }
                            ),
                            new Text($"Период: с {startDate.ToString("dd.MM.yyyy")} по {endDate.ToString("dd.MM.yyyy")}")
                        )
                    );
                    body.AppendChild(dateParagraph);

                    // Empty paragraph for spacing
                    body.AppendChild(new Paragraph(new Run(new Text(""))));

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
                        new TableJustification() { Val = TableRowAlignmentValues.Center },
                        new TableWidth() { Width = "100%", Type = TableWidthUnitValues.Pct }
                    );
                    table.AppendChild(tableProperties);

                    TableRow headerRow = new TableRow();
                    foreach (DataGridViewColumn column in dataGridView1.Columns)
                    {
                        TableCell cell = new TableCell(
                            new TableCellProperties(
                                new Shading() { Fill = "4472C4" }
                            ),
                            new Paragraph(
                                new ParagraphProperties(
                                    new SpacingBetweenLines() { After = "0" }
                                ),
                                new Run(
                                    new RunProperties(
                                        new Bold(),
                                        new Color() { Val = "FFFFFF" },
                                        new FontSize() { Val = "26" }
                                    ),
                                    new Text(column.HeaderText)
                                )
                            )
                        );
                        headerRow.AppendChild(cell);
                    }
                    table.AppendChild(headerRow);

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            TableRow tableRow = new TableRow();
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                string cellValue = cell.Value?.ToString() ?? string.Empty;

                                if (dataGridView1.Columns[cell.ColumnIndex].HeaderText.Contains("Дата") &&
                                    DateTime.TryParse(cellValue, out DateTime dateValue))
                                {
                                    cellValue = dateValue.ToString("dd.MM.yyyy");
                                }

                                TableCell tableCell = new TableCell(
                                    new TableCellProperties(
                                        new Shading() { Fill = row.Index % 2 == 0 ? "E9E9E9" : "FFFFFF" }
                                    ),
                                    new Paragraph(
                                        new ParagraphProperties(
                                            new SpacingBetweenLines() { After = "0" }
                                        ),
                                        new Run(
                                            new RunProperties(
                                                new FontSize() { Val = "26" }
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

                    // Empty paragraph for spacing
                    body.AppendChild(new Paragraph(new Run(new Text(""))));

                    Paragraph totalParagraph = new Paragraph(
                        new ParagraphProperties(
                            new SpacingBetweenLines() { After = "0" }
                        ),
                        new Run(
                            new RunProperties(
                                new Bold(),
                                new FontSize() { Val = "26" }
                            ),
                            new Text($"Общая сумма: {textBox1.Text}")
                        )
                    );
                    body.AppendChild(totalParagraph);

                    mainPart.Document.Save();
                }

                MessageBox.Show("Данные успешно экспортированы в Word.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}");
            }
        }

        private void ExportToExcel(string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Отчет");

                DateTime startDate = dateTimePicker1.Value;
                DateTime endDate = dateTimePicker2.Value;
                string reportTitle;
                switch (comboBox1.Text)
                {
                    case "Поставки":
                        reportTitle = "Отчет по поставкам";
                        break;
                    case "Сотрудники":
                        reportTitle = "Отчет по сотрудникам";
                        break;
                    case "Продажи":
                        reportTitle = "Отчет по продажам";
                        break;
                    default:
                        reportTitle = "Отчет";
                        break;
                }

                // Add title
                worksheet.Cell(1, 1).Value = reportTitle;
                worksheet.Cell(1, 1).Style.Font.Bold = true;
                worksheet.Cell(1, 1).Style.Font.FontSize = 14;
                worksheet.Range(1, 1, 1, dataGridView1.Columns.Count).Merge();

                // Add date period
                worksheet.Cell(2, 1).Value = $"Период: с {startDate.ToString("dd.MM.yyyy")} по {endDate.ToString("dd.MM.yyyy")}";
                worksheet.Cell(2, 1).Style.Font.FontSize = 12;
                worksheet.Range(2, 1, 2, dataGridView1.Columns.Count).Merge();

                // Empty row
                worksheet.Cell(3, 1).Value = string.Empty;
                worksheet.Range(3, 1, 3, dataGridView1.Columns.Count).Merge();

                // Headers start at row 4
                int headerRow = 4;
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cell(headerRow, i + 1).Value = dataGridView1.Columns[i].HeaderText;
                    worksheet.Cell(headerRow, i + 1).Style.Font.Bold = true;
                }

                // Data starts at row 5
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1.Columns[j].HeaderText.Contains("Дата"))
                        {
                            DateTime dateValue;
                            if (DateTime.TryParse(dataGridView1.Rows[i].Cells[j].Value?.ToString(), out dateValue))
                            {
                                worksheet.Cell(i + headerRow + 1, j + 1).Value = dateValue.ToString("dd.MM.yyyy");
                            }
                        }
                        else
                        {
                            worksheet.Cell(i + headerRow + 1, j + 1).Value = dataGridView1.Rows[i].Cells[j].Value?.ToString() ?? string.Empty;
                        }
                    }
                }

                // Empty row
                int lastRow = dataGridView1.Rows.Count + headerRow + 2;
                worksheet.Cell(lastRow, 1).Value = string.Empty;
                worksheet.Range(lastRow, 1, lastRow, dataGridView1.Columns.Count).Merge();

                // Total sum
                worksheet.Cell(lastRow + 1, 1).Value = "Общая сумма:";
                worksheet.Cell(lastRow + 1, 1).Style.Font.Bold = true;
                worksheet.Cell(lastRow + 1, 2).Value = textBox1.Text;

                worksheet.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.RangeUsed().Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                worksheet.Columns().AdjustToContents();

                workbook.SaveAs(filePath);
                MessageBox.Show("Данные успешно экспортированы в Excel.");
            }
        }

        private void экспортДанныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word Documents (*.docx)|*.docx|Excel Files (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "Экспорт данных";

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

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            programinfo form2 = new programinfo();
            form2.Show();
        }

        private async Task button3_ClickAsync(object sender, EventArgs e)
        {
            try
            {
                if (reportData == null)
                {
                    MessageBox.Show("Нет данных для поиска. Сначала сформируйте отчет.");
                    return;
                }

                string searchText = textBox2.Text.Trim().ToLower();

                if (string.IsNullOrEmpty(searchText))
                {
                    dataGridView1.DataSource = reportData;
                    return;
                }

                Cursor.Current = Cursors.WaitCursor;

                DataView filteredView = await Task.Run(() =>
                {
                    string filter = BuildFilter(searchText);
                    DataView dv = new DataView(reportData);
                    dv.RowFilter = filter;
                    return dv;
                });

                dataGridView1.DataSource = filteredView;
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
        private string BuildFilter(string searchText)
        {
            string filter = "";
            foreach (DataColumn column in reportData.Columns)
            {
                if (filter.Length > 0) filter += " OR ";
                filter += $"Convert([{column.ColumnName}], 'System.String') LIKE '%{searchText}%'";
            }
            return filter;
        }


        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            string currentText = textBox2.Text.Trim().ToLower();

            if (currentText == _lastSearchText)
                return;

            _lastSearchText = currentText;
            _searchTimer?.Dispose();
            _searchTimer = new System.Threading.Timer(_ =>
            {
                this.Invoke((MethodInvoker)delegate
                {
                    UpdateAutoComplete(currentText);
                });
            }, null, 300, System.Threading.Timeout.Infinite);
        }

        private void UpdateAutoComplete(string searchText)
        {
            if (dataGridView1.DataSource == null || dataGridView1.Rows.Count == 0)
                return;

            if (string.IsNullOrEmpty(searchText))
            {
                textBox2.AutoCompleteCustomSource = null;
                return;
            }

            var suggestions = new HashSet<string>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null)
                    {
                        string cellValue = cell.Value.ToString();
                        if (cellValue.ToLower().Contains(searchText))
                        {
                            suggestions.Add(cellValue);
                        }
                    }
                }
            }

            if (suggestions.Count > 0)
            {
                var autoComplete = new AutoCompleteStringCollection();
                autoComplete.AddRange(suggestions.ToArray());

                int cursorPos = textBox2.SelectionStart;

                textBox2.AutoCompleteCustomSource = autoComplete;
                textBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                textBox2.AutoCompleteSource = AutoCompleteSource.CustomSource;
                textBox2.SelectionStart = cursorPos;
            }
            else
            {
                textBox2.AutoCompleteCustomSource = null;
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                button3_ClickAsync(sender, e);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (reportData == null)
            {
                MessageBox.Show("Нет данных для поиска. Сначала сформируйте отчет.");
                return;
            }

            string searchText = textBox2.Text.Trim().ToLower();

            if (string.IsNullOrEmpty(searchText))
            {
                dataGridView1.DataSource = reportData;
                return;
            }

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                string filter = "";
                foreach (DataColumn column in reportData.Columns)
                {
                    if (filter.Length > 0) filter += " OR ";
                    filter += $"Convert([{column.ColumnName}], 'System.String') LIKE '%{searchText}%'";
                }

                DataView dv = new DataView(reportData);
                dv.RowFilter = filter;
                dataGridView1.DataSource = dv;
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
            string helpFileName = "Справка формы руководителя.pdf";

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