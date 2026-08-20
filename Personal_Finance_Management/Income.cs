using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Collections.Generic;

namespace Personal_Finance_Management
{
    public partial class Income : UserControl
    {
        private GroupBox groupBox1;
        private Button Exit;
        private Button Addexpenses;
        private TextBox Frequency_income;
        private TextBox Amount_income;
        private TextBox Type_income;
        private TextBox Source_income;
        private Label label5;
        private Label label4;
        private Label label3;
        private Label label2;
        private Label label1;
        private GroupBox groupBox2;
        private Label label6;
        private GroupBox groupBox4;
        private GroupBox groupBox3;
        private DataGridView dataGridView2;
        private DateTimePicker Next_payment_income;
        private TextBox Income_total;
        private Button button1;
        private Button button2;
        private Panel panel1;
        private GroupBox Total;
        private Label label8;
        private Label label9;
        private GroupBox Reports;
        private Label label10;
        private DateTimePicker dateTimePicker2;
        private DateTimePicker dateTimePicker1;
        private Label label7;
        private Button View_reports;
        private DataGridView dataGridView1;
       
        private DataGridViewTextBoxColumn Next_payment;
        private DataGridViewTextBoxColumn Sources;
        private DataGridViewTextBoxColumn types;
        private DataGridViewTextBoxColumn Amount;
        private DataGridViewTextBoxColumn Frequency;
        private DataGridViewTextBoxColumn SN;
        private DataGridViewTextBoxColumn Typess;
        private DataGridViewTextBoxColumn Amounts;
        private Button View_all;
        private TextBox search;

        // Add Update button
        private Button UpdateIncome;

        // Add variable to store selected row ID for update
        private int selectedRowId = -1;

        // Add label for search
        private Label label11;

        string connectionString = AppData.ConnectionString;
        public Income()
        {
            InitializeComponent();
            CreateResponsiveIncome();
        }

        private void CreateResponsiveIncome()
        {
            this.BackColor = Color.FromArgb(240, 240, 245);
            this.Padding = new Padding(20);
            this.Dock = DockStyle.Fill;

            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // ✅ Center all text (cells + headers)
            dataGridView2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

            // Add double-click event handler
            dataGridView2.CellDoubleClick += DataGridView2_CellDoubleClick;

            LoadIncomeFromDatabase();
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // ✅ Center all text (cells + headers)
            dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

            LoadIncomeTypeFromDatabase();
        }

        // Double-click event handler for dataGridView2
        private void DataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && !dataGridView2.Rows[e.RowIndex].IsNewRow)
            {
                DataGridViewRow selectedRow = dataGridView2.Rows[e.RowIndex];

                // Store the selected row ID for update
                selectedRowId = Convert.ToInt32(selectedRow.Tag);

                // Display data in the input fields
                Next_payment_income.Value = DateTime.Parse(selectedRow.Cells["Next_payment"].Value.ToString());
                Source_income.Text = selectedRow.Cells["Sources"].Value.ToString();
                Type_income.Text = selectedRow.Cells["types"].Value.ToString();

                // Remove currency symbols and parse amount
                string amountText = selectedRow.Cells["Amount"].Value.ToString().Replace("$", "").Replace(",", "");
                Amount_income.Text = amountText;

                Frequency_income.Text = selectedRow.Cells["Frequency"].Value.ToString();

                // Enable update button and disable add button
                UpdateIncome.Enabled = true;
                Addexpenses.Enabled = false;
            }
        }

        // Update button click event handler
        private void UpdateIncome_Click(object sender, EventArgs e)
        {
            if (selectedRowId == -1)
            {
                MessageBox.Show("Please select a row to update by double-clicking it.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string Source = Source_income.Text.Trim();
            string Type = Type_income.Text.Trim();
            string AmountText = Amount_income.Text.Trim();
            string Frequency = Frequency_income.Text.Trim();
            string Next_payment = Next_payment_income.Value.ToString("yyyy-MM-dd");

            // Validate inputs
            if (string.IsNullOrWhiteSpace(Source) || string.IsNullOrWhiteSpace(Type) ||
                string.IsNullOrWhiteSpace(AmountText) || string.IsNullOrWhiteSpace(Frequency))
            {
                MessageBox.Show("Please fill in all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(AmountText, out decimal amount))
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
                    string updateQuery = @"UPDATE Income_sources 
                                          SET Source = @Source, 
                                              Type = @Type, 
                                              Amount = @Amount, 
                                              Frequency = @Frequency, 
                                              Next_payment = @Next_payment 
                                          WHERE Id = @Id";

                    using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", selectedRowId);
                        cmd.Parameters.AddWithValue("@Source", Source);
                        cmd.Parameters.AddWithValue("@Type", Type);
                        cmd.Parameters.AddWithValue("@Amount", AmountText);
                        cmd.Parameters.AddWithValue("@Frequency", Frequency);
                        cmd.Parameters.AddWithValue("@Next_payment", Next_payment);

                        int rows = cmd.ExecuteNonQuery();

                        if (rows > 0)
                        {
                            MessageBox.Show($"Income updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            // Clear fields and reset buttons
                            ClearInputFields();
                            UpdateIncome.Enabled = false;
                            Addexpenses.Enabled = true;
                            selectedRowId = -1;

                            // Refresh data
                            LoadIncomeFromDatabase();
                            LoadIncomeTypeFromDatabase();
                        }
                        else
                        {
                            MessageBox.Show("Failed to update income.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            Source_income.Clear();
            Frequency_income.Clear();
            Type_income.Clear();
            Amount_income.Clear();
            Next_payment_income.Value = DateTime.Now;
        }

        // Search functionality
        private void Search_TextChanged(object sender, EventArgs e)
        {
            LoadIncomeFromDatabase(search.Text.Trim());
        }

        private void LoadIncomeFromDatabase(string searchText = "")
        {
            string query = @"
        SELECT Id, Source, Type, Amount, Frequency, Next_payment
        FROM Income_sources
        WHERE (@search = '' OR Source LIKE '%' + @search + '%' OR Type LIKE '%' + @search + '%')
          AND MONTH(Next_payment) = MONTH(GETDATE())
          AND YEAR(Next_payment) = YEAR(GETDATE())
        ORDER BY Next_payment DESC;
    ";

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@search", searchText ?? string.Empty);

                    connection.Open();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        dataGridView2.Rows.Clear();

                        while (reader.Read())
                        {
                            try
                            {
                                string source = reader["Source"]?.ToString() ?? "";
                                string type = reader["Type"]?.ToString() ?? "";
                                decimal amount = reader["Amount"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Amount"])
                                    : 0m;
                                string frequency = reader["Frequency"]?.ToString() ?? "";
                                string nextPayment = reader["Next_payment"] != DBNull.Value
                                    ? Convert.ToDateTime(reader["Next_payment"]).ToString("yyyy-MM-dd")
                                    : "";

                                string formattedAmount = amount.ToString("C");

                                // Add row (without showing Id)
                                int rowIndex = dataGridView2.Rows.Add(nextPayment, source, type, formattedAmount, frequency);

                                // Store the Id in the Tag property
                                dataGridView2.Rows[rowIndex].Tag = reader["Id"];
                            }
                            catch (Exception rowEx)
                            {
                                MessageBox.Show($"Error reading income record: {rowEx.Message}",
                                                "Row Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
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

        private void LoadIncomeTypeFromDatabase()
        {
            try
            {
                // Dictionary to hold total amounts per type
                Dictionary<string, decimal> totalsByType = new Dictionary<string, decimal>();

                // Iterate through each row of dataGridView2
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.IsNewRow) continue; // Skip the new row placeholder

                    string type = row.Cells[2].Value?.ToString() ?? ""; // Assuming "Type" is at index 2
                    string amountStr = row.Cells[3].Value?.ToString() ?? "$0"; // Assuming "Amount" is at index 3

                    // Remove currency symbol and parse
                    if (decimal.TryParse(amountStr, System.Globalization.NumberStyles.Currency, null, out decimal amount))
                    {
                        if (totalsByType.ContainsKey(type))
                            totalsByType[type] += amount;
                        else
                            totalsByType[type] = amount;
                    }
                }

                // Populate dataGridView1
                dataGridView1.Rows.Clear();
                int sn = 1;
                foreach (var kvp in totalsByType)
                {
                    dataGridView1.Rows.Add(sn, kvp.Key, kvp.Value.ToString("C"));
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

                foreach (DataGridViewRow row in dataGridView2.Rows)
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
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            Label label = new Label();
            label.Text = labelText;
            label.Font = new Font("Segoe UI", 10);
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;

            control.Dock = DockStyle.Fill;
            control.Font = new Font("Segoe UI", 10);
            control.Margin = new Padding(0, 5, 0, 5);

            layout.Controls.Add(label, 0, row);
            layout.Controls.Add(control, 1, row);
        }

        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.UpdateIncome = new System.Windows.Forms.Button();
            this.Next_payment_income = new System.Windows.Forms.DateTimePicker();
            this.Exit = new System.Windows.Forms.Button();
            this.Addexpenses = new System.Windows.Forms.Button();
            this.Frequency_income = new System.Windows.Forms.TextBox();
            this.Amount_income = new System.Windows.Forms.TextBox();
            this.Type_income = new System.Windows.Forms.TextBox();
            this.Source_income = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label11 = new System.Windows.Forms.Label();
            this.search = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.SN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Typess = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Amounts = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label7 = new System.Windows.Forms.Label();
            this.Reports = new System.Windows.Forms.GroupBox();
            this.View_all = new System.Windows.Forms.Button();
            this.View_reports = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.label9 = new System.Windows.Forms.Label();
            this.Total = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.Income_total = new System.Windows.Forms.TextBox();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.Next_payment = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Sources = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.types = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Amount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Frequency = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.Reports.SuspendLayout();
            this.Total.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.UpdateIncome);
            this.groupBox1.Controls.Add(this.Next_payment_income);
            this.groupBox1.Controls.Add(this.Exit);
            this.groupBox1.Controls.Add(this.Addexpenses);
            this.groupBox1.Controls.Add(this.Frequency_income);
            this.groupBox1.Controls.Add(this.Amount_income);
            this.groupBox1.Controls.Add(this.Type_income);
            this.groupBox1.Controls.Add(this.Source_income);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(0, 580);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(875, 188);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Add/Update Income";
            // 
            // UpdateIncome
            // 
            this.UpdateIncome.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.UpdateIncome.BackColor = System.Drawing.Color.DodgerBlue;
            this.UpdateIncome.Enabled = false;
            this.UpdateIncome.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.UpdateIncome.Location = new System.Drawing.Point(545, 56);
            this.UpdateIncome.Name = "UpdateIncome";
            this.UpdateIncome.Size = new System.Drawing.Size(143, 55);
            this.UpdateIncome.TabIndex = 12;
            this.UpdateIncome.Text = "Update Income";
            this.UpdateIncome.UseVisualStyleBackColor = false;
            this.UpdateIncome.Click += new System.EventHandler(this.UpdateIncome_Click);
            // 
            // Next_payment_income
            // 
            this.Next_payment_income.Font = new System.Drawing.Font("Segoe UI", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Next_payment_income.Location = new System.Drawing.Point(232, 30);
            this.Next_payment_income.Name = "Next_payment_income";
            this.Next_payment_income.Size = new System.Drawing.Size(254, 31);
            this.Next_payment_income.TabIndex = 0;
            this.Next_payment_income.Value = new System.DateTime(2025, 11, 2, 14, 54, 39, 0);
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
            this.Addexpenses.Size = new System.Drawing.Size(143, 55);
            this.Addexpenses.TabIndex = 10;
            this.Addexpenses.Text = "Add Income";
            this.Addexpenses.UseVisualStyleBackColor = false;
            this.Addexpenses.Click += new System.EventHandler(this.Addexpenses_Click);
            // 
            // Frequency_income
            // 
            this.Frequency_income.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Frequency_income.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Frequency_income.Location = new System.Drawing.Point(232, 152);
            this.Frequency_income.Name = "Frequency_income";
            this.Frequency_income.Size = new System.Drawing.Size(254, 30);
            this.Frequency_income.TabIndex = 8;
            // 
            // Amount_income
            // 
            this.Amount_income.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Amount_income.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Amount_income.Location = new System.Drawing.Point(232, 122);
            this.Amount_income.Name = "Amount_income";
            this.Amount_income.Size = new System.Drawing.Size(254, 30);
            this.Amount_income.TabIndex = 7;
            // 
            // Type_income
            // 
            this.Type_income.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Type_income.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Type_income.Location = new System.Drawing.Point(232, 92);
            this.Type_income.Name = "Type_income";
            this.Type_income.Size = new System.Drawing.Size(254, 30);
            this.Type_income.TabIndex = 6;
            // 
            // Source_income
            // 
            this.Source_income.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Source_income.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Source_income.Location = new System.Drawing.Point(232, 62);
            this.Source_income.Name = "Source_income";
            this.Source_income.Size = new System.Drawing.Size(254, 30);
            this.Source_income.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(59, 33);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(95, 28);
            this.label5.TabIndex = 4;
            this.label5.Text = "Paid Date";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(56, 147);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(106, 28);
            this.label4.TabIndex = 3;
            this.label4.Text = "Frequency:";
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
            this.label2.Size = new System.Drawing.Size(57, 28);
            this.label2.TabIndex = 1;
            this.label2.Text = "Type:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(56, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 28);
            this.label1.TabIndex = 0;
            this.label1.Text = "Source:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.search);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.groupBox4);
            this.groupBox2.Controls.Add(this.groupBox3);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(875, 580);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Income Details";
            // 
            // label11
            // 
            this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(274, 39);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(126, 23);
            this.label11.TabIndex = 16;
            this.label11.Text = "Search by Text";
            // 
            // search
            // 
            this.search.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.search.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.search.Location = new System.Drawing.Point(406, 36);
            this.search.Name = "search";
            this.search.Size = new System.Drawing.Size(133, 27);
            this.search.TabIndex = 15;
            this.search.TextChanged += new System.EventHandler(this.Search_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(6, 24);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(237, 41);
            this.label6.TabIndex = 2;
            this.label6.Text = "Recent Incomes";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.dataGridView1);
            this.groupBox4.Controls.Add(this.label7);
            this.groupBox4.Controls.Add(this.Reports);
            this.groupBox4.Controls.Add(this.panel1);
            this.groupBox4.Dock = System.Windows.Forms.DockStyle.Right;
            this.groupBox4.Location = new System.Drawing.Point(545, 18);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(327, 559);
            this.groupBox4.TabIndex = 1;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Income Types";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeColumns = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.SN,
            this.Typess,
            this.Amounts});
            this.dataGridView1.Location = new System.Drawing.Point(9, 49);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(315, 212);
            this.dataGridView1.TabIndex = 12;
            // 
            // SN
            // 
            this.SN.FillWeight = 64.17112F;
            this.SN.HeaderText = "SN";
            this.SN.MinimumWidth = 6;
            this.SN.Name = "SN";
            // 
            // Typess
            // 
            this.Typess.FillWeight = 117.9144F;
            this.Typess.HeaderText = "Types";
            this.Typess.MinimumWidth = 6;
            this.Typess.Name = "Typess";
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
            this.label7.Location = new System.Drawing.Point(105, 18);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(118, 23);
            this.label7.TabIndex = 11;
            this.label7.Text = "Income Types";
            // 
            // Reports
            // 
            this.Reports.Controls.Add(this.View_all);
            this.Reports.Controls.Add(this.View_reports);
            this.Reports.Controls.Add(this.label10);
            this.Reports.Controls.Add(this.dateTimePicker2);
            this.Reports.Controls.Add(this.label9);
            this.Reports.Controls.Add(this.Total);
            this.Reports.Controls.Add(this.dateTimePicker1);
            this.Reports.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.Reports.Location = new System.Drawing.Point(3, 302);
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
            // dateTimePicker2
            // 
            this.dateTimePicker2.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dateTimePicker2.Location = new System.Drawing.Point(154, 48);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(154, 27);
            this.dateTimePicker2.TabIndex = 11;
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
            this.label8.Location = new System.Drawing.Point(15, 26);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(112, 23);
            this.label8.TabIndex = 9;
            this.label8.Text = "Total Income";
            // 
            // Income_total
            // 
            this.Income_total.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.Income_total.Location = new System.Drawing.Point(133, 23);
            this.Income_total.Name = "Income_total";
            this.Income_total.Size = new System.Drawing.Size(166, 30);
            this.Income_total.TabIndex = 3;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dateTimePicker1.Location = new System.Drawing.Point(154, 18);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(154, 27);
            this.dateTimePicker1.TabIndex = 9;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(3, 500);
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
            this.groupBox3.Controls.Add(this.dataGridView2);
            this.groupBox3.Location = new System.Drawing.Point(6, 73);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(533, 492);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Income list";
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.AllowUserToDeleteRows = false;
            this.dataGridView2.AllowUserToResizeColumns = false;
            this.dataGridView2.AllowUserToResizeRows = false;
            this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView2.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Next_payment,
            this.Sources,
            this.types,
            this.Amount,
            this.Frequency});
            this.dataGridView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView2.GridColor = System.Drawing.SystemColors.ButtonShadow;
            this.dataGridView2.Location = new System.Drawing.Point(3, 18);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowHeadersVisible = false;
            this.dataGridView2.RowHeadersWidth = 51;
            this.dataGridView2.RowTemplate.Height = 24;
            this.dataGridView2.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.dataGridView2.Size = new System.Drawing.Size(527, 471);
            this.dataGridView2.TabIndex = 0;
            // 
            // Next_payment
            // 
            this.Next_payment.HeaderText = "Paid Date";
            this.Next_payment.MinimumWidth = 6;
            this.Next_payment.Name = "Next_payment";
            // 
            // Sources
            // 
            this.Sources.HeaderText = "Sources";
            this.Sources.MinimumWidth = 6;
            this.Sources.Name = "Sources";
            // 
            // types
            // 
            this.types.HeaderText = "Types";
            this.types.MinimumWidth = 6;
            this.types.Name = "types";
            // 
            // Amount
            // 
            this.Amount.HeaderText = "Amount";
            this.Amount.MinimumWidth = 6;
            this.Amount.Name = "Amount";
            // 
            // Frequency
            // 
            this.Frequency.HeaderText = "Frequency";
            this.Frequency.MinimumWidth = 6;
            this.Frequency.Name = "Frequency";
            // 
            // Income
            // 
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Income";
            this.Size = new System.Drawing.Size(875, 768);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.Reports.ResumeLayout(false);
            this.Reports.PerformLayout();
            this.Total.ResumeLayout(false);
            this.Total.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.ResumeLayout(false);

        }

        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Addexpenses_Click(object sender, EventArgs e)
        {
            string Source = Source_income.Text.Trim();
            string Type = Type_income.Text.Trim();
            string AmountText = Amount_income.Text.Trim();
            string Frequency = Frequency_income.Text.Trim();
            string Next_payment = Next_payment_income.Value.ToString("yyyy-MM-dd");

            // Validate inputs
            if (string.IsNullOrWhiteSpace(Source) || string.IsNullOrWhiteSpace(Type) ||
                string.IsNullOrWhiteSpace(AmountText) || string.IsNullOrWhiteSpace(Frequency))
            {
                MessageBox.Show("Please fill in all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(AmountText, out decimal amount))
            {
                MessageBox.Show("Please enter a valid amount.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // Get next ID
                    int nextId = 1;
                    string idQuery = "SELECT ISNULL(MAX(Id), 0) + 1 FROM Income_sources";
                    int nextId1 = 1;
                    string idQuery1 = "SELECT ISNULL(MAX(Id), 0) + 1 FROM Recent_transaction";

                    using (SqlCommand idCmd = new SqlCommand(idQuery, conn))
                    {
                        nextId = Convert.ToInt32(idCmd.ExecuteScalar());
                    }
                    using (SqlCommand idCmd = new SqlCommand(idQuery1, conn))
                    {
                        nextId1 = Convert.ToInt32(idCmd.ExecuteScalar());
                    }

                    // Insert new income
                    string insertQuery = @"INSERT INTO Income_sources (Id, Source, Type, Amount, Frequency, Next_payment)
                                           VALUES (@Id, @Source, @Type, @Amount, @Frequency, @Next_payment)";

                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", nextId);
                        cmd.Parameters.AddWithValue("@Source", Source);
                        cmd.Parameters.AddWithValue("@Type", Type);
                        cmd.Parameters.AddWithValue("@Amount", AmountText);
                        cmd.Parameters.AddWithValue("@Frequency", Frequency);
                        cmd.Parameters.AddWithValue("@Next_payment", Next_payment);

                        int rows = cmd.ExecuteNonQuery();

                        if (rows > 0)
                        {
                            MessageBox.Show($"Income added successfully with ID: {nextId}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            // Clear input fields
                            ClearInputFields();
                            LoadIncomeFromDatabase();
                            LoadIncomeTypeFromDatabase();
                        }
                        else
                        {
                            MessageBox.Show("Failed to add income.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    string insertRecentQuery = @"INSERT INTO Recent_transaction (Id, Date, Description, Category, Amount )
                                         VALUES (@Id, @Date, @Description, @Category, @Amount )";
                    using (SqlCommand cmdRecent = new SqlCommand(insertRecentQuery, conn))
                    {
                        cmdRecent.Parameters.AddWithValue("@Id", nextId1);
                        cmdRecent.Parameters.AddWithValue("@Date", Next_payment);
                        cmdRecent.Parameters.AddWithValue("@Description", Source);
                        cmdRecent.Parameters.AddWithValue("@Category", Type);
                        cmdRecent.Parameters.AddWithValue("@Amount", AmountText);

                        cmdRecent.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void View_reports_Click(object sender, EventArgs e)
        {
            DateTime fromDate = dateTimePicker1.Value.Date;
            DateTime toDate = dateTimePicker2.Value.Date;

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
            Source, 
            Type, 
            Amount, 
            Frequency, 
            Next_payment
        FROM Income_sources
        WHERE TRY_CONVERT(date, Next_payment) BETWEEN @FromDate AND @ToDate
        ORDER BY TRY_CONVERT(date, Next_payment) DESC;
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
                        dataGridView2.Rows.Clear();

                        while (reader.Read())
                        {
                            try
                            {
                                string source = reader["Source"]?.ToString() ?? "";
                                string type = reader["Type"]?.ToString() ?? "";
                                decimal amount = reader["Amount"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Amount"])
                                    : 0m;
                                string frequency = reader["Frequency"]?.ToString() ?? "";
                                string nextPaymentStr = reader["Next_payment"]?.ToString();
                                string formattedNextPayment = "";

                                if (DateTime.TryParse(nextPaymentStr, out DateTime nextPaymentDate))
                                    formattedNextPayment = nextPaymentDate.ToString("yyyy-MM-dd");

                                string formattedAmount = amount.ToString("C");

                                // Add to DataGridView
                                dataGridView2.Rows.Add(formattedNextPayment, source, type, formattedAmount, frequency);
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
                LoadIncomeTypeFromDatabase();
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

        private void button1_Click(object sender, EventArgs e)
        {
            // If no data, warn user
            if (dataGridView2.Rows.Count == 0)
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
                saveFileDialog.FileName = "Income_Report_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // Build CSV content
                        StringBuilder csvContent = new StringBuilder();

                        // Write headers
                        string[] columnNames = dataGridView2.Columns
                            .Cast<DataGridViewColumn>()
                            .Select(col => "\"" + col.HeaderText.Replace("\"", "\"\"") + "\"")
                            .ToArray();
                        csvContent.AppendLine(string.Join(",", columnNames));

                        // Write rows
                        foreach (DataGridViewRow row in dataGridView2.Rows)
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
            if (dataGridView2.SelectedRows.Count == 0)
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
                DataGridViewRow selectedRow = dataGridView2.SelectedRows[0];

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

                    string deleteQuery = "DELETE FROM Income_sources WHERE Id = @Id";
                    using (SqlCommand cmd = new SqlCommand(deleteQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", idToDelete);
                        int rows = cmd.ExecuteNonQuery();

                        if (rows > 0)
                        {
                            // Remove row from DataGridView
                            dataGridView2.Rows.Remove(selectedRow);

                            // Update total amount
                            UpdateTotalAmountInTextBox();
                            LoadIncomeTypeFromDatabase();

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

        private void View_all_Click(object sender, EventArgs e)
        {
            LoadIncomeFromDatabase();
            LoadIncomeTypeFromDatabase();
        }
        public void RefreshData()
        {
            LoadIncomeFromDatabase();
            LoadIncomeTypeFromDatabase();
        }
    }
}