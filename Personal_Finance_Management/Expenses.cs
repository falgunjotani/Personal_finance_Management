using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace Personal_Finance_Management
{
    public partial class Expenses : UserControl
    {
        private GroupBox groupBox1;
        private DateTimePicker Expense_date;
        private Button Exit;
        private Button Addexpenses;
        private TextBox Expense_paymethod;
        private TextBox Expenses_amount;
        private TextBox Expense_cat;
        private TextBox Expense_des;
        private Label label5;
        private Label label4;
        private Label label3;
        private Label label2;
        private Label label1;
        private GroupBox groupBox2;
        private Label label6;
        private GroupBox groupBox4;
        private DataGridView dataGridView11;
        private Label label7;
        private GroupBox Reports;
        private Button View_all;
        private Button View_reports;
        private Label label10;
        private DateTimePicker dateTimePicker21;
        private Label label9;
        private GroupBox Total;
        private Label label8;
        private TextBox Income_total;
        private DateTimePicker dateTimePicker11;
        private Panel panel1;
        private Button button1;
        private Button button2;
        private GroupBox groupBox3;
        private DataGridView dataGridView21;
        private DataGridViewTextBoxColumn SN;
        private DataGridViewTextBoxColumn Categories;
        private DataGridViewTextBoxColumn Amounts;
        private DataGridViewTextBoxColumn Date;
        private DataGridViewTextBoxColumn Description;
        private DataGridViewTextBoxColumn Category;
        private DataGridViewTextBoxColumn Amount;
        private DataGridViewTextBoxColumn Payment_Method;
      
        private Label label11;
        private TextBox Search;

        // Add Update button
        private Button UpdateExpense;

        // Add variable to store selected row ID for update
        private int selectedRowId = -1;

        string connectionString = AppData.ConnectionString;

        public Expenses()
        {
            InitializeComponent();
            CreateResponsiveExpenses();
        }

        private void CreateResponsiveExpenses()
        {
            this.BackColor = Color.FromArgb(240, 240, 245);
            this.Padding = new Padding(20);
            this.Dock = DockStyle.Fill;
            dataGridView21.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView21.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // ✅ Center all text (cells + headers)
            dataGridView21.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView21.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView21.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            dataGridView21.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

            // Add double-click event handler
            dataGridView21.CellDoubleClick += DataGridView21_CellDoubleClick;

            LoadExpensesFromDatabase();
            dataGridView11.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView11.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // ✅ Center all text (cells + headers)
            dataGridView11.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView11.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView11.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            dataGridView11.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            LoadexpensescategoryFromDatabase();
        }

        // Double-click event handler for dataGridView21
        private void DataGridView21_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && !dataGridView21.Rows[e.RowIndex].IsNewRow)
            {
                DataGridViewRow selectedRow = dataGridView21.Rows[e.RowIndex];

                // Store the selected row ID for update
                selectedRowId = Convert.ToInt32(selectedRow.Tag);

                // Display data in the input fields
                Expense_date.Value = DateTime.Parse(selectedRow.Cells["Date"].Value.ToString());
                Expense_des.Text = selectedRow.Cells["Description"].Value.ToString();
                Expense_cat.Text = selectedRow.Cells["Category"].Value.ToString();

                // Remove currency symbols and parse amount
                string amountText = selectedRow.Cells["Amount"].Value.ToString().Replace("$", "").Replace(",", "");
                Expenses_amount.Text = amountText;

                Expense_paymethod.Text = selectedRow.Cells["Payment_Method"].Value.ToString();

                // Enable update button and disable add button
                UpdateExpense.Enabled = true;
                Addexpenses.Enabled = false;
            }
        }

        // Update button click event handler
        private void UpdateExpense_Click(object sender, EventArgs e)
        {
            if (selectedRowId == -1)
            {
                MessageBox.Show("Please select a row to update by double-clicking it.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string date = Expense_date.Value.ToString("yyyy-MM-dd");
            string amountText = Expenses_amount.Text.Trim();
            string description = Expense_des.Text.Trim();
            string paymentMethod = Expense_paymethod.Text.Trim();
            string category = Expense_cat.Text.Trim();

            // Validate inputs
            if (string.IsNullOrWhiteSpace(date) || string.IsNullOrWhiteSpace(description) ||
                string.IsNullOrWhiteSpace(category) || string.IsNullOrWhiteSpace(amountText) ||
                string.IsNullOrWhiteSpace(paymentMethod))
            {
                MessageBox.Show("Please fill in all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(amountText, out decimal amount))
            {
                MessageBox.Show("Please enter a valid amount.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // Update query
                    string updateQuery = @"UPDATE Recent_expenses 
                                          SET Date = @Date, 
                                              Description = @Description, 
                                              Category = @Category, 
                                              Amount = @Amount, 
                                              Payment_method = @Payment_method 
                                          WHERE Id = @Id";

                    using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", selectedRowId);
                        cmd.Parameters.AddWithValue("@Date", date);
                        cmd.Parameters.AddWithValue("@Description", description);
                        cmd.Parameters.AddWithValue("@Category", category);
                        cmd.Parameters.AddWithValue("@Amount", amount);
                        cmd.Parameters.AddWithValue("@Payment_method", paymentMethod);

                        int rows = cmd.ExecuteNonQuery();

                        if (rows > 0)
                        {
                            MessageBox.Show($"Expense updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            // Clear fields and reset buttons
                            ClearInputFields();
                            UpdateExpense.Enabled = false;
                            Addexpenses.Enabled = true;
                            selectedRowId = -1;

                            // Refresh data
                            LoadExpensesFromDatabase();
                            LoadexpensescategoryFromDatabase();
                        }
                        else
                        {
                            MessageBox.Show("Failed to update expense.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Method to clear input fields
        private void ClearInputFields()
        {
            Expenses_amount.Clear();
            Expense_des.Clear();
            Expense_paymethod.Clear();
            Expense_cat.Clear();
            Expense_date.Value = DateTime.Now;
        }

        private void LoadExpensesFromDatabase1212(string searchText = "")
        {
            string query = @"
        SELECT Id, Date, Description, Category, Amount, Payment_method
        FROM Recent_expenses
        WHERE(@search = '' OR Asset LIKE '%' + @search + '%') AND MONTH(Date) = MONTH(GETDATE())
          AND YEAR(Date) = YEAR(GETDATE())
        ORDER BY Date DESC;
    ";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    connection.Open();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        dataGridView21.Rows.Clear();

                        while (reader.Read())
                        {
                            try
                            {
                                string date = reader["Date"] != DBNull.Value
                                    ? Convert.ToDateTime(reader["Date"]).ToString("yyyy-MM-dd")
                                    : "";
                                string description = reader["Description"]?.ToString() ?? "";
                                string category = reader["Category"]?.ToString() ?? "";
                                decimal amount = reader["Amount"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Amount"])
                                    : 0m;
                                string payment_m = reader["Payment_method"]?.ToString() ?? "";

                                string formattedAmount = amount.ToString("C");

                                // Add row (without showing Id)
                                int rowIndex = dataGridView21.Rows.Add(date, description, category, formattedAmount, payment_m);

                                // Store the Id in the Tag property
                                dataGridView21.Rows[rowIndex].Tag = reader["Id"];
                            }
                            catch (Exception rowEx)
                            {
                                MessageBox.Show($"Error reading expense record: {rowEx.Message}",
                                                "Row Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(
                    $"Database error: {ex.Message}\n\nPlease make sure:\n1. SQL Server is running\n2. Database 'personal_finance' exists\n3. Table 'Recent_expenses' exists",
                    "Database Connection Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading expense data: {ex.Message}",
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        private void LoadExpensesFromDatabase(string searchText = "")
        {
            string query = @"
    SELECT Id, Date, Description, Category, Amount, Payment_method
    FROM Recent_expenses
    WHERE (@search = '' OR Description LIKE '%' + @search + '%') 
      AND MONTH(Date) = MONTH(GETDATE())
      AND YEAR(Date) = YEAR(GETDATE())
    ORDER BY Date DESC;
";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    // Add the search parameter
                    command.Parameters.AddWithValue("@search", searchText ?? string.Empty);

                    connection.Open();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        dataGridView21.Rows.Clear();

                        while (reader.Read())
                        {
                            try
                            {
                                string date = reader["Date"] != DBNull.Value
                                    ? Convert.ToDateTime(reader["Date"]).ToString("yyyy-MM-dd")
                                    : "";
                                string description = reader["Description"]?.ToString() ?? "";
                                string category = reader["Category"]?.ToString() ?? "";
                                decimal amount = reader["Amount"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Amount"])
                                    : 0m;
                                string payment_m = reader["Payment_method"]?.ToString() ?? "";

                                string formattedAmount = amount.ToString("C");

                                // Add row (without showing Id)
                                int rowIndex = dataGridView21.Rows.Add(date, description, category, formattedAmount, payment_m);

                                // Store the Id in the Tag property
                                dataGridView21.Rows[rowIndex].Tag = reader["Id"];
                            }
                            catch (Exception rowEx)
                            {
                                MessageBox.Show($"Error reading expense record: {rowEx.Message}",
                                                "Row Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(
                    $"Database error: {ex.Message}\n\nPlease make sure:\n1. SQL Server is running\n2. Database 'personal_finance' exists\n3. Table 'Recent_expenses' exists",
                    "Database Connection Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading expense data: {ex.Message}",
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        private void LoadexpensescategoryFromDatabase()
        {
            try
            {
                // Dictionary to hold total amounts per type
                Dictionary<string, decimal> totalsByType = new Dictionary<string, decimal>();

                // Iterate through each row of dataGridView2
                foreach (DataGridViewRow row in dataGridView21.Rows)
                {
                    if (row.IsNewRow) continue; // Skip the new row placeholder

                    string categ = row.Cells[2].Value?.ToString() ?? ""; // Assuming "Type" is at index 2
                    string amountStr = row.Cells[3].Value?.ToString() ?? "$0"; // Assuming "Amount" is at index 3

                    // Remove currency symbol and parse
                    if (decimal.TryParse(amountStr, System.Globalization.NumberStyles.Currency, null, out decimal amount))
                    {
                        if (totalsByType.ContainsKey(categ))
                            totalsByType[categ] += amount;
                        else
                            totalsByType[categ] = amount;
                    }
                }

                // Populate dataGridView1
                dataGridView11.Rows.Clear();
                int sn = 1;
                foreach (var kvp in totalsByType)
                {
                    dataGridView11.Rows.Add(sn, kvp.Key, kvp.Value.ToString("C"));
                    sn++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error summarizing income data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            UpdateTotalAmountInTextBox();
        }

        private void UpdateTotalAmountInTextBox()
        {
            try
            {
                decimal totalAmount = 0m;

                foreach (DataGridViewRow row in dataGridView21.Rows)
                {
                    if (row.IsNewRow) continue; // Skip the new row placeholder

                    string amountStr = row.Cells[3].Value?.ToString() ?? "$0"; // Assuming "Amount" is at index 3

                    // Remove currency symbol and parse
                    if (decimal.TryParse(amountStr, System.Globalization.NumberStyles.Currency, null, out decimal amount))
                    {
                        totalAmount += amount;
                    }
                }

                // Display the total in your TextBox
                Income_total.Text = totalAmount.ToString("C"); // "C" formats as currency
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error calculating total amount: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddFormControl(TableLayoutPanel layout, string labelText, Control control, int row)
        {
            Label label = new Label();
            label.Text = labelText;
            label.Font = new Font("Segoe UI", 10);
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.Margin = new Padding(0, 5, 10, 5);

            control.Dock = DockStyle.Fill;
            control.Margin = new Padding(0, 5, 0, 5);
            control.Font = new Font("Segoe UI", 10);

            layout.Controls.Add(label, 0, row);
            layout.Controls.Add(control, 1, row);
        }

        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.UpdateExpense = new System.Windows.Forms.Button();
            this.Expense_date = new System.Windows.Forms.DateTimePicker();
            this.Exit = new System.Windows.Forms.Button();
            this.Addexpenses = new System.Windows.Forms.Button();
            this.Expense_paymethod = new System.Windows.Forms.TextBox();
            this.Expenses_amount = new System.Windows.Forms.TextBox();
            this.Expense_cat = new System.Windows.Forms.TextBox();
            this.Expense_des = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label11 = new System.Windows.Forms.Label();
            this.Search = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.dataGridView11 = new System.Windows.Forms.DataGridView();
            this.SN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Categories = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Amounts = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label7 = new System.Windows.Forms.Label();
            this.Reports = new System.Windows.Forms.GroupBox();
            this.View_all = new System.Windows.Forms.Button();
            this.View_reports = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.dateTimePicker21 = new System.Windows.Forms.DateTimePicker();
            this.label9 = new System.Windows.Forms.Label();
            this.Total = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.Income_total = new System.Windows.Forms.TextBox();
            this.dateTimePicker11 = new System.Windows.Forms.DateTimePicker();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dataGridView21 = new System.Windows.Forms.DataGridView();
            this.Date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Description = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Category = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Amount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Payment_Method = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView11)).BeginInit();
            this.Reports.SuspendLayout();
            this.Total.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView21)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.UpdateExpense);
            this.groupBox1.Controls.Add(this.Expense_date);
            this.groupBox1.Controls.Add(this.Exit);
            this.groupBox1.Controls.Add(this.Addexpenses);
            this.groupBox1.Controls.Add(this.Expense_paymethod);
            this.groupBox1.Controls.Add(this.Expenses_amount);
            this.groupBox1.Controls.Add(this.Expense_cat);
            this.groupBox1.Controls.Add(this.Expense_des);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(0, 592);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(875, 188);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Add/Update Expenses";
            // 
            // UpdateExpense
            // 
            this.UpdateExpense.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.UpdateExpense.BackColor = System.Drawing.Color.DodgerBlue;
            this.UpdateExpense.Enabled = false;
            this.UpdateExpense.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.UpdateExpense.Location = new System.Drawing.Point(545, 56);
            this.UpdateExpense.Name = "UpdateExpense";
            this.UpdateExpense.Size = new System.Drawing.Size(165, 55);
            this.UpdateExpense.TabIndex = 12;
            this.UpdateExpense.Text = "Update Expense";
            this.UpdateExpense.UseVisualStyleBackColor = false;
            this.UpdateExpense.Click += new System.EventHandler(this.UpdateExpense_Click);
            // 
            // Expense_date
            // 
            this.Expense_date.Font = new System.Drawing.Font("Segoe UI", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Expense_date.Location = new System.Drawing.Point(232, 30);
            this.Expense_date.Name = "Expense_date";
            this.Expense_date.Size = new System.Drawing.Size(254, 31);
            this.Expense_date.TabIndex = 0;
            this.Expense_date.Value = new System.DateTime(2025, 11, 2, 14, 54, 39, 0);
            // 
            // Exit
            // 
            this.Exit.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.Exit.BackColor = System.Drawing.Color.OrangeRed;
            this.Exit.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.Exit.Location = new System.Drawing.Point(716, 117);
            this.Exit.Name = "Exit";
            this.Exit.Size = new System.Drawing.Size(134, 55);
            this.Exit.TabIndex = 11;
            this.Exit.Text = "Exit";
            this.Exit.UseVisualStyleBackColor = false;
            this.Exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // Addexpenses
            // 
            this.Addexpenses.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.Addexpenses.BackColor = System.Drawing.Color.LimeGreen;
            this.Addexpenses.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.Addexpenses.Location = new System.Drawing.Point(545, 117);
            this.Addexpenses.Name = "Addexpenses";
            this.Addexpenses.Size = new System.Drawing.Size(165, 55);
            this.Addexpenses.TabIndex = 10;
            this.Addexpenses.Text = "Add Expenses";
            this.Addexpenses.UseVisualStyleBackColor = false;
            this.Addexpenses.Click += new System.EventHandler(this.Addexpenses_Click);
            // 
            // Expense_paymethod
            // 
            this.Expense_paymethod.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Expense_paymethod.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Expense_paymethod.Location = new System.Drawing.Point(232, 152);
            this.Expense_paymethod.Name = "Expense_paymethod";
            this.Expense_paymethod.Size = new System.Drawing.Size(254, 30);
            this.Expense_paymethod.TabIndex = 8;
            // 
            // Expenses_amount
            // 
            this.Expenses_amount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Expenses_amount.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Expenses_amount.Location = new System.Drawing.Point(232, 122);
            this.Expenses_amount.Name = "Expenses_amount";
            this.Expenses_amount.Size = new System.Drawing.Size(254, 30);
            this.Expenses_amount.TabIndex = 7;
            // 
            // Expense_cat
            // 
            this.Expense_cat.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Expense_cat.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Expense_cat.Location = new System.Drawing.Point(232, 92);
            this.Expense_cat.Name = "Expense_cat";
            this.Expense_cat.Size = new System.Drawing.Size(254, 30);
            this.Expense_cat.TabIndex = 6;
            // 
            // Expense_des
            // 
            this.Expense_des.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Expense_des.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Expense_des.Location = new System.Drawing.Point(232, 62);
            this.Expense_des.Name = "Expense_des";
            this.Expense_des.Size = new System.Drawing.Size(254, 30);
            this.Expense_des.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(59, 33);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 28);
            this.label5.TabIndex = 4;
            this.label5.Text = "Date";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(56, 147);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(166, 28);
            this.label4.TabIndex = 3;
            this.label4.Text = "Payment Method:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(56, 119);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(87, 28);
            this.label3.TabIndex = 2;
            this.label3.Text = "Amount:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(56, 91);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(96, 28);
            this.label2.TabIndex = 1;
            this.label2.Text = "Category:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(56, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 28);
            this.label1.TabIndex = 0;
            this.label1.Text = "Description:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.Search);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.groupBox4);
            this.groupBox2.Controls.Add(this.groupBox3);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(875, 592);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Expenses Details";
            // 
            // label11
            // 
            this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(240, 47);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(167, 23);
            this.label11.TabIndex = 6;
            this.label11.Text = "Search by Category";
            // 
            // Search
            // 
            this.Search.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.Search.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Search.Location = new System.Drawing.Point(403, 43);
            this.Search.Name = "Search";
            this.Search.Size = new System.Drawing.Size(137, 27);
            this.Search.TabIndex = 5;
            this.Search.TextChanged += new System.EventHandler(this.Search_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(6, 29);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(249, 41);
            this.label6.TabIndex = 2;
            this.label6.Text = "Recent Expenses";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.dataGridView11);
            this.groupBox4.Controls.Add(this.label7);
            this.groupBox4.Controls.Add(this.Reports);
            this.groupBox4.Controls.Add(this.panel1);
            this.groupBox4.Dock = System.Windows.Forms.DockStyle.Right;
            this.groupBox4.Location = new System.Drawing.Point(545, 18);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(327, 571);
            this.groupBox4.TabIndex = 1;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Expenses Categories";
            // 
            // dataGridView11
            // 
            this.dataGridView11.AllowUserToAddRows = false;
            this.dataGridView11.AllowUserToDeleteRows = false;
            this.dataGridView11.AllowUserToResizeColumns = false;
            this.dataGridView11.AllowUserToResizeRows = false;
            this.dataGridView11.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView11.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView11.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView11.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView11.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.SN,
            this.Categories,
            this.Amounts});
            this.dataGridView11.Location = new System.Drawing.Point(9, 49);
            this.dataGridView11.Name = "dataGridView11";
            this.dataGridView11.RowHeadersVisible = false;
            this.dataGridView11.RowHeadersWidth = 51;
            this.dataGridView11.RowTemplate.Height = 24;
            this.dataGridView11.Size = new System.Drawing.Size(315, 212);
            this.dataGridView11.TabIndex = 12;
            // 
            // SN
            // 
            this.SN.FillWeight = 64.17112F;
            this.SN.HeaderText = "SN";
            this.SN.MinimumWidth = 6;
            this.SN.Name = "SN";
            // 
            // Categories
            // 
            this.Categories.FillWeight = 117.9144F;
            this.Categories.HeaderText = "Category";
            this.Categories.MinimumWidth = 6;
            this.Categories.Name = "Categories";
            // 
            // Amounts
            // 
            this.Amounts.FillWeight = 117.9144F;
            this.Amounts.HeaderText = "Amount";
            this.Amounts.MinimumWidth = 6;
            this.Amounts.Name = "Amounts";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold);
            this.label7.Location = new System.Drawing.Point(67, 23);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(182, 23);
            this.label7.TabIndex = 11;
            this.label7.Text = "Expenses by category";
            // 
            // Reports
            // 
            this.Reports.Controls.Add(this.View_all);
            this.Reports.Controls.Add(this.View_reports);
            this.Reports.Controls.Add(this.label10);
            this.Reports.Controls.Add(this.dateTimePicker21);
            this.Reports.Controls.Add(this.label9);
            this.Reports.Controls.Add(this.Total);
            this.Reports.Controls.Add(this.dateTimePicker11);
            this.Reports.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.Reports.Location = new System.Drawing.Point(3, 314);
            this.Reports.Name = "Reports";
            this.Reports.Size = new System.Drawing.Size(321, 198);
            this.Reports.TabIndex = 10;
            this.Reports.TabStop = false;
            this.Reports.Text = "Report";
            // 
            // View_all
            // 
            this.View_all.BackColor = System.Drawing.Color.LightYellow;
            this.View_all.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.View_all.Location = new System.Drawing.Point(162, 75);
            this.View_all.Name = "View_all";
            this.View_all.Size = new System.Drawing.Size(156, 35);
            this.View_all.TabIndex = 12;
            this.View_all.Text = "📊 View All";
            this.View_all.UseVisualStyleBackColor = false;
            this.View_all.Click += new System.EventHandler(this.View_all_Click);
            // 
            // View_reports
            // 
            this.View_reports.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.View_reports.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.View_reports.Location = new System.Drawing.Point(6, 75);
            this.View_reports.Name = "View_reports";
            this.View_reports.Size = new System.Drawing.Size(156, 35);
            this.View_reports.TabIndex = 6;
            this.View_reports.Text = "📊 View Reports";
            this.View_reports.UseVisualStyleBackColor = false;
            this.View_reports.Click += new System.EventHandler(this.View_reports_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold);
            this.label10.Location = new System.Drawing.Point(73, 22);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(52, 23);
            this.label10.TabIndex = 9;
            this.label10.Text = "From";
            // 
            // dateTimePicker21
            // 
            this.dateTimePicker21.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dateTimePicker21.Location = new System.Drawing.Point(154, 48);
            this.dateTimePicker21.Name = "dateTimePicker21";
            this.dateTimePicker21.Size = new System.Drawing.Size(154, 27);
            this.dateTimePicker21.TabIndex = 11;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold);
            this.label9.Location = new System.Drawing.Point(73, 49);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(28, 23);
            this.label9.TabIndex = 10;
            this.label9.Text = "To";
            // 
            // Total
            // 
            this.Total.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Total.Controls.Add(this.label8);
            this.Total.Controls.Add(this.Income_total);
            this.Total.Location = new System.Drawing.Point(3, 120);
            this.Total.Name = "Total";
            this.Total.Size = new System.Drawing.Size(312, 72);
            this.Total.TabIndex = 8;
            this.Total.TabStop = false;
            this.Total.Text = "Total";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(3, 26);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(125, 23);
            this.label8.TabIndex = 9;
            this.label8.Text = "Total Expenses";
            // 
            // Income_total
            // 
            this.Income_total.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.Income_total.Location = new System.Drawing.Point(133, 23);
            this.Income_total.Name = "Income_total";
            this.Income_total.Size = new System.Drawing.Size(166, 30);
            this.Income_total.TabIndex = 3;
            // 
            // dateTimePicker11
            // 
            this.dateTimePicker11.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dateTimePicker11.Location = new System.Drawing.Point(154, 18);
            this.dateTimePicker11.Name = "dateTimePicker11";
            this.dateTimePicker11.Size = new System.Drawing.Size(154, 27);
            this.dateTimePicker11.TabIndex = 9;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(3, 512);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(321, 56);
            this.panel1.TabIndex = 6;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button1.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.button1.Location = new System.Drawing.Point(16, 11);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(145, 35);
            this.button1.TabIndex = 4;
            this.button1.Text = "📤 Export Data";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.LightCyan;
            this.button2.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.button2.Location = new System.Drawing.Point(167, 11);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(145, 35);
            this.button2.TabIndex = 5;
            this.button2.Text = "📤 Delete Data";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.Controls.Add(this.dataGridView21);
            this.groupBox3.Location = new System.Drawing.Point(6, 73);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(533, 504);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Expenses list";
            // 
            // dataGridView21
            // 
            this.dataGridView21.AllowUserToAddRows = false;
            this.dataGridView21.AllowUserToDeleteRows = false;
            this.dataGridView21.AllowUserToResizeColumns = false;
            this.dataGridView21.AllowUserToResizeRows = false;
            this.dataGridView21.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView21.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dataGridView21.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView21.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Date,
            this.Description,
            this.Category,
            this.Amount,
            this.Payment_Method});
            this.dataGridView21.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView21.GridColor = System.Drawing.SystemColors.ButtonShadow;
            this.dataGridView21.Location = new System.Drawing.Point(3, 18);
            this.dataGridView21.Name = "dataGridView21";
            this.dataGridView21.RowHeadersVisible = false;
            this.dataGridView21.RowHeadersWidth = 51;
            this.dataGridView21.RowTemplate.Height = 24;
            this.dataGridView21.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dataGridView21.Size = new System.Drawing.Size(527, 483);
            this.dataGridView21.TabIndex = 0;
            // 
            // Date
            // 
            this.Date.HeaderText = "Date";
            this.Date.MinimumWidth = 6;
            this.Date.Name = "Date";
            // 
            // Description
            // 
            this.Description.HeaderText = "Description";
            this.Description.MinimumWidth = 6;
            this.Description.Name = "Description";
            // 
            // Category
            // 
            this.Category.HeaderText = "Category";
            this.Category.MinimumWidth = 6;
            this.Category.Name = "Category";
            // 
            // Amount
            // 
            this.Amount.HeaderText = "Amount";
            this.Amount.MinimumWidth = 6;
            this.Amount.Name = "Amount";
            // 
            // Payment_Method
            // 
            this.Payment_Method.HeaderText = "Pay Method";
            this.Payment_Method.MinimumWidth = 6;
            this.Payment_Method.Name = "Payment_Method";
            // 
            // Expenses
            // 
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Expenses";
            this.Size = new System.Drawing.Size(875, 780);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView11)).EndInit();
            this.Reports.ResumeLayout(false);
            this.Reports.PerformLayout();
            this.Total.ResumeLayout(false);
            this.Total.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView21)).EndInit();
            this.ResumeLayout(false);
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void Addexpenses_Click(object sender, EventArgs e)
        {
            string date = Expense_date.Value.ToString("yyyy-MM-dd");
            string amountText = Expenses_amount.Text.Trim();
            string description = Expense_des.Text.Trim();
            string paymentMethod = Expense_paymethod.Text.Trim();
            string category = Expense_cat.Text.Trim();

            // Validate inputs
            if (string.IsNullOrWhiteSpace(date) || string.IsNullOrWhiteSpace(description) ||
                string.IsNullOrWhiteSpace(category) || string.IsNullOrWhiteSpace(amountText) ||
                string.IsNullOrWhiteSpace(paymentMethod))
            {
                MessageBox.Show("Please fill in all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(amountText, out decimal amount))
            {
                MessageBox.Show("Please enter a valid amount.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // Get next ID for Recent_expenses
                    int nextId = 1;
                    string idQuery = "SELECT ISNULL(MAX(Id), 0) + 1 FROM Recent_expenses";

                    using (SqlCommand idCmd = new SqlCommand(idQuery, conn))
                    {
                        nextId = Convert.ToInt32(idCmd.ExecuteScalar());
                    }

                    // Insert new expense
                    string insertQuery = @"INSERT INTO Recent_expenses (Id, Date, Description, Category, Amount, Payment_method)
                                   VALUES (@Id, @Date, @Description, @Category, @Amount, @Payment_method)";

                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", nextId);
                        cmd.Parameters.AddWithValue("@Date", date);
                        cmd.Parameters.AddWithValue("@Description", description);
                        cmd.Parameters.AddWithValue("@Category", category);
                        cmd.Parameters.AddWithValue("@Amount", amount);
                        cmd.Parameters.AddWithValue("@Payment_method", paymentMethod);

                        int rows = cmd.ExecuteNonQuery();

                        if (rows > 0)
                        {
                            // Check if category exists in Budget table and update spent amount
                            UpdateBudgetSpent(conn, category, amount, date);

                            MessageBox.Show($"Expense added successfully with ID: {nextId}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            // Clear input fields
                            ClearInputFields();
                            LoadExpensesFromDatabase();
                            LoadexpensescategoryFromDatabase();
                        }
                        else
                        {
                            MessageBox.Show("Failed to add expense.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Method to update budget spent amount
        private void UpdateBudgetSpent(SqlConnection conn, string category, decimal amount, string date)
        {
            try
            {
                // First, check if the category exists in the Budget table for the current month
                string checkQuery = @"
            SELECT Id, Spent, Remaining 
            FROM Budget 
            WHERE Category = @Category 
            AND MONTH(Date) = MONTH(@Date) 
            AND YEAR(Date) = YEAR(@Date)";

                using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                {
                    checkCmd.Parameters.AddWithValue("@Category", category);
                    checkCmd.Parameters.AddWithValue("@Date", date);

                    using (SqlDataReader reader = checkCmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            // Category exists - update the spent amount
                            int budgetId = Convert.ToInt32(reader["Id"]);

                            // Safe conversion for Spent
                            decimal currentSpent = 0m;
                            if (reader["Spent"] != DBNull.Value)
                            {
                                currentSpent = Convert.ToDecimal(reader["Spent"]);
                            }

                            // Safe conversion for Remaining
                            decimal currentRemaining = 0m;
                            if (reader["Remaining"] != DBNull.Value)
                            {
                                currentRemaining = Convert.ToDecimal(reader["Remaining"]);
                            }

                            reader.Close(); // Close the reader before executing another command

                            // Only update Remaining (add the expense amount to current remaining)
                            decimal newRemaining = currentRemaining + amount;

                            string updateQuery = @"
                        UPDATE Budget 
                        SET Remaining = @NewRemaining
                        WHERE Id = @BudgetId";

                            using (SqlCommand updateCmd = new SqlCommand(updateQuery, conn))
                            {
                                updateCmd.Parameters.AddWithValue("@NewRemaining", newRemaining);
                                updateCmd.Parameters.AddWithValue("@BudgetId", budgetId);
                                updateCmd.ExecuteNonQuery();
                            }

                            Console.WriteLine($"Budget updated: Category={category}, NewRemaining={newRemaining}");
                        }
                        else
                        {
                            reader.Close(); // Category doesn't exist in Budget table
                            Console.WriteLine($"No budget found for category: {category} in current month");
                            // You can choose to ignore or create a new entry
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the detailed error but don't stop the expense addition process
                Console.WriteLine($"Budget update error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");

                MessageBox.Show($"Note: Expense was added but budget update failed: {ex.Message}",
                                "Budget Update Warning",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
            }
        }
        //private void Addexpenses_Click(object sender, EventArgs e)
        //{
        //    string date = Expense_date.Value.ToString("yyyy-MM-dd");
        //    string amountText = Expenses_amount.Text.Trim();
        //    string description = Expense_des.Text.Trim();
        //    string paymentMethod = Expense_paymethod.Text.Trim();
        //    string category = Expense_cat.Text.Trim();

        //    // Validate inputs
        //    if (string.IsNullOrWhiteSpace(date) || string.IsNullOrWhiteSpace(description) ||
        //        string.IsNullOrWhiteSpace(category) || string.IsNullOrWhiteSpace(amountText) ||
        //        string.IsNullOrWhiteSpace(paymentMethod))
        //    {
        //        MessageBox.Show("Please fill in all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    if (!decimal.TryParse(amountText, out decimal amount))
        //    {
        //        MessageBox.Show("Please enter a valid amount.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        return;
        //    }

        //    try
        //    {
        //        using (SqlConnection conn = new SqlConnection(connectionString))
        //        {
        //            conn.Open();

        //            // Get next ID
        //            int nextId = 1;
        //            string idQuery = "SELECT ISNULL(MAX(Id), 0) + 1 FROM Recent_expenses";

        //            using (SqlCommand idCmd = new SqlCommand(idQuery, conn))
        //            {
        //                nextId = Convert.ToInt32(idCmd.ExecuteScalar());
        //            }

        //            // Insert new expense
        //            string insertQuery = @"INSERT INTO Recent_expenses (Id, Date, Description, Category, Amount, Payment_method)
        //                                   VALUES (@Id, @Date, @Description, @Category, @Amount, @Payment_method)";

        //            using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
        //            {
        //                cmd.Parameters.AddWithValue("@Id", nextId);
        //                cmd.Parameters.AddWithValue("@Date", date);
        //                cmd.Parameters.AddWithValue("@Description", description);
        //                cmd.Parameters.AddWithValue("@Category", category);
        //                cmd.Parameters.AddWithValue("@Amount", amount);
        //                cmd.Parameters.AddWithValue("@Payment_method", paymentMethod);

        //                int rows = cmd.ExecuteNonQuery();

        //                if (rows > 0)
        //                {
        //                    MessageBox.Show($"Expense added successfully with ID: {nextId}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

        //                    // Clear input fields
        //                    ClearInputFields();
        //                    LoadExpensesFromDatabase();
        //                    LoadexpensescategoryFromDatabase();
        //                }
        //                else
        //                {
        //                    MessageBox.Show("Failed to add expense.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

        private void Exit_Click_(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void View_reports_Click(object sender, EventArgs e)
        {
            DateTime fromDate = dateTimePicker11.Value.Date;
            DateTime toDate = dateTimePicker21.Value.Date;

            // Validate date range
            if (fromDate > toDate)
            {
                MessageBox.Show("The 'From' date cannot be after the 'To' date.",
                                "Invalid Date Range",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                return;
            }

            string query = @"
        SELECT 
            Date, 
            Description, 
            Category, 
            Amount, 
            Payment_method
        FROM Recent_expenses
        WHERE TRY_CONVERT(date, Date) BETWEEN @FromDate AND @ToDate
        ORDER BY TRY_CONVERT(date, Date) DESC;
    ";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.Add("@FromDate", SqlDbType.Date).Value = fromDate;
                    command.Parameters.Add("@ToDate", SqlDbType.Date).Value = toDate;

                    connection.Open();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        dataGridView21.Rows.Clear();

                        while (reader.Read())
                        {
                            try
                            {
                                string description = reader["Description"]?.ToString() ?? "";
                                string category = reader["Category"]?.ToString() ?? "";
                                decimal amount = reader["Amount"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Amount"])
                                    : 0m;
                                string payment_m = reader["Payment_method"]?.ToString() ?? "";
                                string nextPaymentStr = reader["Date"]?.ToString();
                                string formattedNextPayment = "";

                                if (DateTime.TryParse(nextPaymentStr, out DateTime nextPaymentDate))
                                    formattedNextPayment = nextPaymentDate.ToString("yyyy-MM-dd");

                                string formattedAmount = amount.ToString("C");

                                // Add to DataGridView
                                dataGridView21.Rows.Add(formattedNextPayment, description, category, formattedAmount, payment_m);
                            }
                            catch (Exception rowEx)
                            {
                                MessageBox.Show($"Error reading income record: {rowEx.Message}",
                                                "Row Error",
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
                LoadexpensescategoryFromDatabase();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(
                    $"Database error: {ex.Message}\n\nPlease make sure:\n1. SQL Server is running\n2. Database 'personal_finance' exists\n3. Table 'Income_sources' exists",
                    "Database Connection Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading income data: {ex.Message}",
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        private void View_all_Click(object sender, EventArgs e)
        {
            LoadExpensesFromDatabase();
            LoadexpensescategoryFromDatabase();
        }

        public void RefreshData()
        {
            // Reload data from database
            LoadExpensesFromDatabase(); // or whatever your load method is called
            LoadexpensescategoryFromDatabase(); // or other refresh methods
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            // If no data, warn user
            if (dataGridView21.Rows.Count == 0)
            {
                MessageBox.Show("No data available to export.",
                                "Export to CSV",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                return;
            }

            // Let user choose where to save file
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "CSV files (*.csv)|*.csv";
                saveFileDialog.Title = "Save Report as CSV";
                saveFileDialog.FileName = "Expenses_Report_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Build CSV content
                        StringBuilder csvContent = new StringBuilder();

                        // Write headers
                        string[] columnNames = dataGridView21.Columns
                            .Cast<DataGridViewColumn>()
                            .Select(col => "\"" + col.HeaderText.Replace("\"", "\"\"") + "\"")
                            .ToArray();
                        csvContent.AppendLine(string.Join(",", columnNames));

                        // Write rows
                        foreach (DataGridViewRow row in dataGridView21.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                string[] cells = row.Cells
                                    .Cast<DataGridViewCell>()
                                    .Select(cell => "\"" + (cell.Value?.ToString().Replace("\"", "\"\"") ?? "") + "\"")
                                    .ToArray();
                                csvContent.AppendLine(string.Join(",", cells));
                            }
                        }

                        // Save file
                        File.WriteAllText(saveFileDialog.FileName, csvContent.ToString(), Encoding.UTF8);

                        MessageBox.Show("CSV file saved successfully:\n" + saveFileDialog.FileName,
                                        "Export Complete",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error exporting CSV: " + ex.Message,
                                        "Export Error",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Check if a row is selected
            if (dataGridView21.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a row to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Confirm deletion
            DialogResult result = MessageBox.Show(
                "Are you sure you want to delete the selected data?",
                "Confirm Deletion",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question
            );

            if (result != DialogResult.OK)
                return; // User canceled

            try
            {
                // Get the selected row
                DataGridViewRow selectedRow = dataGridView21.SelectedRows[0];

                // Retrieve the Id from the Tag property
                if (selectedRow.Tag == null)
                {
                    MessageBox.Show("Cannot determine the record to delete.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int idToDelete = Convert.ToInt32(selectedRow.Tag);

                // Delete from database
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string deleteQuery = "DELETE FROM Recent_expenses WHERE Id = @Id";
                    using (SqlCommand cmd = new SqlCommand(deleteQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", idToDelete);
                        int rows = cmd.ExecuteNonQuery();

                        if (rows > 0)
                        {
                            // Remove row from DataGridView
                            dataGridView21.Rows.Remove(selectedRow);

                            // Update total amount
                            UpdateTotalAmountInTextBox();
                            LoadexpensescategoryFromDatabase();

                            MessageBox.Show("Record deleted successfully.", "Deleted", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("No record found to delete.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error deleting data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Search_TextChanged(object sender, EventArgs e)
        {
            LoadExpensesFromDatabase(Search.Text.Trim());
        }
    }
}