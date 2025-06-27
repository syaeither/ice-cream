using Npgsql;
using System;
using System.Data;
using System.Windows.Forms;

namespace ice_cream
{
    public partial class history_client : Form
    {
        private int clientId;
        private string connectionString = Program.ConnectionString;
        public history_client(int clientId)
        {
            InitializeComponent();
            this.clientId = clientId;
            LoadOrderHistory();
        }

        private void LoadOrderHistory()
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
    o.total_cost AS ""Сумма"",
    o.payment_method AS ""Способ оплаты"",
    e.full_name AS ""Продавец""
FROM 
    orders o
LEFT JOIN 
    employees e ON o.code_employee = e.id_employee
WHERE 
    o.code_client = @clientId
ORDER BY 
    o.date_orders DESC;";

                    using (var command = new NpgsqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@clientId", clientId);
                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                DataTable dt = new DataTable();
                                dt.Columns.Add("Дата заказа", typeof(DateTime));
                                dt.Columns.Add("Товары", typeof(string));
                                dt.Columns.Add("Сумма", typeof(decimal));
                                dt.Columns.Add("Способ оплаты", typeof(string));
                                dt.Columns.Add("Продавец", typeof(string));

                                while (reader.Read())
                                {
                                    DateTime orderDate = reader.GetDateTime(0);
                                    string products = reader.IsDBNull(1) ? "Нет данных" : reader.GetString(1);
                                    decimal totalSum = reader.GetDecimal(2);
                                    string paymentMethod = reader.IsDBNull(3) ? "Не указан" : reader.GetString(3);
                                    string employeeName = reader.IsDBNull(4) ? "Не указан" : reader.GetString(4);

                                    dt.Rows.Add(orderDate, products, totalSum, paymentMethod, employeeName);
                                }

                                dataGridView1.DataSource = dt;
                                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                            }
                            else
                            {
                                MessageBox.Show("У этого клиента нет истории заказов");
                                this.Close();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при загрузке истории заказов: {ex.Message}");
                }
                dataGridView1.AllowUserToAddRows = false;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}