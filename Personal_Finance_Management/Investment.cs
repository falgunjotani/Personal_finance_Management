using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Personal_Finance_Management
{
    public partial class Investment : UserControl
    {
        private GroupBox groupBox1;
        private DateTimePicker Purchase_date;
        private Button Exit;
        private Button Addinvestment;
        private TextBox Share_units;
        private TextBox Amount_invest;
        private TextBox Purchase_price;
        private TextBox Purchase_asset;
        private Label label5;
        private Label label4;
        private Label label3;
        private Label label2;
        private Label label1;
        private GroupBox groupBox2;
        private Label label6;
        private GroupBox groupBox4;
        private DataGridView dataGridView1;
        private Label label7;
        private GroupBox Reports;
        private Button View_all;
        private Button View_reports;
        private Label label10;
        private DateTimePicker dateTimePicker2;
        private Label label9;
        private GroupBox Total;
        private Label label8;
        private TextBox Income_total;
        private DateTimePicker dateTimePicker1;
        private Panel panel1;
        private Button button1;
        private Button button2;
        private GroupBox groupBox3;
        private DataGridView dataGridView21;
        private DataGridViewTextBoxColumn date;
        private DataGridViewTextBoxColumn asset;
        private DataGridViewTextBoxColumn types;
        private DataGridViewTextBoxColumn Amount_invested;
        private DataGridViewTextBoxColumn Share_unit;
       
        private DataGridViewTextBoxColumn SN;
        private DataGridViewTextBoxColumn Typess;
        private DataGridViewTextBoxColumn Amounts;
        private TextBox Search;
        private Label label11;

        // Add Update button
        private Button UpdateInvestment;

        // Add variable to store selected row ID for update
        private int selectedRowId = -1;
        string connectionString = AppData.ConnectionString;
        

        public Investment()
        {
            InitializeComponent();
            CreateResponsiveInvestment();
        }
        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Check if the user is closing the form (not programmatically)
            if (e.CloseReason == CloseReason.UserClosing)
            {
                Application.Exit(); // Exit the entire application
            }
        }
        private void CreateResponsiveInvestment()
        {
            dataGridView21.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView21.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // ✅ Center all text (cells + headers)
            dataGridView21.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView21.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView21.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            dataGridView21.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

            // Add double-click event handler
            dataGridView21.CellDoubleClick += DataGridView21_CellDoubleClick;

            LoadportfolioFromDatabase();
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // ✅ Center all text (cells + headers)
            dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            fractionalshare();
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
                Purchase_date.Value = DateTime.Parse(selectedRow.Cells["date"].Value.ToString());
                Purchase_asset.Text = selectedRow.Cells["asset"].Value.ToString();

                // Remove currency symbols and parse values - SIMPLE VERSION
                string purchasePriceText = selectedRow.Cells["types"].Value.ToString();
                string amountInvestText = selectedRow.Cells["Amount_invested"].Value.ToString();

                // Remove common currency symbols and formatting
                Purchase_price.Text = CleanCurrencyString(purchasePriceText);
                Amount_invest.Text = CleanCurrencyString(amountInvestText);

                Share_units.Text = selectedRow.Cells["Share_unit"].Value.ToString();

                // Enable update button and disable add button
                UpdateInvestment.Enabled = true;
               // Addinvestment.Enabled = false;
            }
        }

        // Simple helper method to remove currency formatting
        private string CleanCurrencyString(string currencyText)
        {
            if (string.IsNullOrEmpty(currencyText))
                return "0";

            // Remove common currency symbols and commas
            return currencyText.Replace("¥", "").Replace("$", "").Replace("€", "").Replace("£", "").Replace(",", "").Trim();
        }

        // Update button click event handler
        private void UpdateInvestment_Click(object sender, EventArgs e)
        {
            if (selectedRowId == -1)
            {
                MessageBox.Show("Please select a row to update by double-clicking it.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string date = Purchase_date.Value.ToString("yyyy-MM-dd");
            string Purchaseasset = Purchase_asset.Text.Trim();
            string Purchaseprice = Purchase_price.Text.Trim();
            string Amountinvest = Amount_invest.Text.Trim();
            string Shareunits = Share_units.Text.Trim();

            // Validate inputs
            if (string.IsNullOrWhiteSpace(date) || string.IsNullOrWhiteSpace(Purchaseasset) ||
                string.IsNullOrWhiteSpace(Purchaseprice) || string.IsNullOrWhiteSpace(Amountinvest) ||
                string.IsNullOrWhiteSpace(Shareunits))
            {
                MessageBox.Show("Please fill in all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(Amountinvest, out decimal amount))
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
                    string updateQuery = @"UPDATE Portfolio 
                                          SET Asset = @Asset, 
                                              Purchase_price = @Purchase_price, 
                                              Amount_invested = @Amount_invested, 
                                              Share_unit = @Share_unit, 
                                              Date = @Date 
                                          WHERE Id = @Id";

                    using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", selectedRowId);
                        cmd.Parameters.AddWithValue("@Asset", Purchaseasset);
                        cmd.Parameters.AddWithValue("@Purchase_price", Purchaseprice);
                        cmd.Parameters.AddWithValue("@Amount_invested", Amountinvest);
                        cmd.Parameters.AddWithValue("@Share_unit", Shareunits);
                        cmd.Parameters.AddWithValue("@Date", date);

                        int rows = cmd.ExecuteNonQuery();

                        if (rows > 0)
                        {
                            MessageBox.Show($"Investment updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            // Clear fields and reset buttons
                            ClearInputFields();
                            UpdateInvestment.Enabled = false;
                            Addinvestment.Enabled = true;
                            selectedRowId = -1;

                            // Refresh data
                            LoadportfolioFromDatabase();
                            fractionalshare();
                        }
                        else
                        {
                            MessageBox.Show("Failed to update investment.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            Purchase_asset.Clear();
            Purchase_price.Clear();
            Amount_invest.Clear();
            Share_units.Clear();
            Purchase_date.Value = DateTime.Now;
        }

        private void fractionalshare()
        {
            try
            {
                // Clear previous rows before populating new data
                dataGridView1.Rows.Clear();

                // Dictionary to store cumulative totals for each asset
                Dictionary<string, decimal> assetTotals = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

                // Iterate through each row in dataGridView21
                foreach (DataGridViewRow row in dataGridView21.Rows)
                {
                    if (row.IsNewRow) continue; // Skip placeholder row

                    // Get Asset name (column index 1)
                    string asset = row.Cells[1].Value?.ToString()?.Trim() ?? "";

                    // Get the progress/share_unit value (column index 4)
                    string progress = row.Cells[4].Value?.ToString()?.Replace("%", "").Trim() ?? "0";

                    // Try to parse the numeric value
                    if (decimal.TryParse(progress, out decimal progressValue))
                    {
                        // Add or update the total for this asset
                        if (assetTotals.ContainsKey(asset))
                            assetTotals[asset] += progressValue;
                        else
                            assetTotals[asset] = progressValue;
                    }
                    else
                    {
                        Console.WriteLine($"Invalid numeric value '{progress}' for asset '{asset}'");
                    }
                }

                // Now display the aggregated results in dataGridView1
                int sn = 1;
                foreach (var kvp in assetTotals)
                {
                    string asset = kvp.Key;
                    decimal totalProgress = kvp.Value;

                    // Add row: S/N, Asset, Total Progress
                    dataGridView1.Rows.Add(sn, asset, totalProgress.ToString("N2")); // format with 2 decimal places
                    sn++;
                }
                UpdateTotalAmountInTextBox();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing fractional shares: {ex.Message}",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.UpdateInvestment = new System.Windows.Forms.Button();
            this.Purchase_date = new System.Windows.Forms.DateTimePicker();
            this.Exit = new System.Windows.Forms.Button();
            this.Addinvestment = new System.Windows.Forms.Button();
            this.Share_units = new System.Windows.Forms.TextBox();
            this.Amount_invest = new System.Windows.Forms.TextBox();
            this.Purchase_price = new System.Windows.Forms.TextBox();
            this.Purchase_asset = new System.Windows.Forms.TextBox();
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
            this.dataGridView21 = new System.Windows.Forms.DataGridView();
            this.date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.asset = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.types = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Amount_invested = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Share_unit = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.Reports.SuspendLayout();
            this.Total.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView21)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.UpdateInvestment);
            this.groupBox1.Controls.Add(this.Purchase_date);
            this.groupBox1.Controls.Add(this.Exit);
            this.groupBox1.Controls.Add(this.Addinvestment);
            this.groupBox1.Controls.Add(this.Share_units);
            this.groupBox1.Controls.Add(this.Amount_invest);
            this.groupBox1.Controls.Add(this.Purchase_price);
            this.groupBox1.Controls.Add(this.Purchase_asset);
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
            this.groupBox1.Text = "Add/Update Investment";
            // 
            // UpdateInvestment
            // 
            this.UpdateInvestment.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.UpdateInvestment.BackColor = System.Drawing.Color.DodgerBlue;
            this.UpdateInvestment.Enabled = false;
            this.UpdateInvestment.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.UpdateInvestment.Location = new System.Drawing.Point(532, 56);
            this.UpdateInvestment.Name = "UpdateInvestment";
            this.UpdateInvestment.Size = new System.Drawing.Size(172, 55);
            this.UpdateInvestment.TabIndex = 12;
            this.UpdateInvestment.Text = "Update Investment";
            this.UpdateInvestment.UseVisualStyleBackColor = false;
            this.UpdateInvestment.Click += new System.EventHandler(this.UpdateInvestment_Click);
            // 
            // Purchase_date
            // 
            this.Purchase_date.Font = new System.Drawing.Font("Segoe UI", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Purchase_date.Location = new System.Drawing.Point(232, 30);
            this.Purchase_date.Name = "Purchase_date";
            this.Purchase_date.Size = new System.Drawing.Size(254, 31);
            this.Purchase_date.TabIndex = 0;
            this.Purchase_date.Value = new System.DateTime(2025, 11, 2, 14, 54, 39, 0);
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
            // Addinvestment
            // 
            this.Addinvestment.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.Addinvestment.BackColor = System.Drawing.Color.LimeGreen;
            this.Addinvestment.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.Addinvestment.Location = new System.Drawing.Point(532, 117);
            this.Addinvestment.Name = "Addinvestment";
            this.Addinvestment.Size = new System.Drawing.Size(172, 55);
            this.Addinvestment.TabIndex = 10;
            this.Addinvestment.Text = "Add Investment";
            this.Addinvestment.UseVisualStyleBackColor = false;
            this.Addinvestment.Click += new System.EventHandler(this.Addinvestment_Click);
            // 
            // Share_units
            // 
            this.Share_units.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Share_units.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Share_units.Location = new System.Drawing.Point(232, 152);
            this.Share_units.Name = "Share_units";
            this.Share_units.Size = new System.Drawing.Size(254, 30);
            this.Share_units.TabIndex = 8;
            // 
            // Amount_invest
            // 
            this.Amount_invest.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Amount_invest.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Amount_invest.Location = new System.Drawing.Point(232, 122);
            this.Amount_invest.Name = "Amount_invest";
            this.Amount_invest.Size = new System.Drawing.Size(254, 30);
            this.Amount_invest.TabIndex = 7;
            // 
            // Purchase_price
            // 
            this.Purchase_price.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Purchase_price.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Purchase_price.Location = new System.Drawing.Point(232, 92);
            this.Purchase_price.Name = "Purchase_price";
            this.Purchase_price.Size = new System.Drawing.Size(254, 30);
            this.Purchase_price.TabIndex = 6;
            // 
            // Purchase_asset
            // 
            this.Purchase_asset.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Purchase_asset.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Purchase_asset.Location = new System.Drawing.Point(232, 62);
            this.Purchase_asset.Name = "Purchase_asset";
            this.Purchase_asset.Size = new System.Drawing.Size(254, 30);
            this.Purchase_asset.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(59, 33);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(135, 28);
            this.label5.TabIndex = 4;
            this.label5.Text = "Purchase Date";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(56, 147);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(115, 28);
            this.label4.TabIndex = 3;
            this.label4.Text = "Share Units:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(56, 119);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(143, 28);
            this.label3.TabIndex = 2;
            this.label3.Text = "Amount Invest:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(56, 91);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(140, 28);
            this.label2.TabIndex = 1;
            this.label2.Text = "Purchase Price:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(59, 64);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 28);
            this.label1.TabIndex = 0;
            this.label1.Text = "Asset:";
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
            this.groupBox2.Text = "Portfolio Details";
            // 
            // label11
            // 
            this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(289, 41);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(135, 23);
            this.label11.TabIndex = 4;
            this.label11.Text = "Search by Asset";
            // 
            // Search
            // 
            this.Search.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.Search.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Search.Location = new System.Drawing.Point(416, 40);
            this.Search.Name = "Search";
            this.Search.Size = new System.Drawing.Size(137, 27);
            this.Search.TabIndex = 3;
            this.Search.TextChanged += new System.EventHandler(this.Search_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(6, 23);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(249, 41);
            this.label6.TabIndex = 2;
            this.label6.Text = "Portfolio Report";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.dataGridView1);
            this.groupBox4.Controls.Add(this.label7);
            this.groupBox4.Controls.Add(this.Reports);
            this.groupBox4.Controls.Add(this.panel1);
            this.groupBox4.Dock = System.Windows.Forms.DockStyle.Right;
            this.groupBox4.Location = new System.Drawing.Point(571, 18);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(301, 571);
            this.groupBox4.TabIndex = 1;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "portfolio Categories";
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
            this.dataGridView1.Size = new System.Drawing.Size(283, 212);
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
            this.Typess.HeaderText = "Asset";
            this.Typess.MinimumWidth = 6;
            this.Typess.Name = "Typess";
            // 
            // Amounts
            // 
            this.Amounts.FillWeight = 117.9144F;
            this.Amounts.HeaderText = "Share Unit";
            this.Amounts.MinimumWidth = 6;
            this.Amounts.Name = "Amounts";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold);
            this.label7.Location = new System.Drawing.Point(44, 23);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(132, 23);
            this.label7.TabIndex = 11;
            this.label7.Text = "Share Holdings";
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
            this.Reports.Location = new System.Drawing.Point(3, 310);
            this.Reports.Name = "Reports";
            this.Reports.Size = new System.Drawing.Size(295, 202);
            this.Reports.TabIndex = 10;
            this.Reports.TabStop = false;
            this.Reports.Text = "Time based Report";
            // 
            // View_all
            // 
            this.View_all.BackColor = System.Drawing.Color.LightYellow;
            this.View_all.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.View_all.Location = new System.Drawing.Point(162, 80);
            this.View_all.Name = "View_all";
            this.View_all.Size = new System.Drawing.Size(127, 35);
            this.View_all.TabIndex = 12;
            this.View_all.Text = "📊 View All";
            this.View_all.UseVisualStyleBackColor = false;
            this.View_all.Click += new System.EventHandler(this.View_all_Click);
            // 
            // View_reports
            // 
            this.View_reports.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.View_reports.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.View_reports.Location = new System.Drawing.Point(6, 80);
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
            this.label10.Location = new System.Drawing.Point(41, 23);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(52, 23);
            this.label10.TabIndex = 9;
            this.label10.Text = "From";
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dateTimePicker2.Location = new System.Drawing.Point(122, 49);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(154, 27);
            this.dateTimePicker2.TabIndex = 11;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold);
            this.label9.Location = new System.Drawing.Point(41, 50);
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
            this.Total.Size = new System.Drawing.Size(286, 72);
            this.Total.TabIndex = 8;
            this.Total.TabStop = false;
            this.Total.Text = "Total";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(-4, 26);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(143, 23);
            this.label8.TabIndex = 9;
            this.label8.Text = "Total Investment";
            // 
            // Income_total
            // 
            this.Income_total.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.Income_total.Location = new System.Drawing.Point(139, 21);
            this.Income_total.Name = "Income_total";
            this.Income_total.Size = new System.Drawing.Size(141, 30);
            this.Income_total.TabIndex = 3;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dateTimePicker1.Location = new System.Drawing.Point(122, 19);
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
            this.panel1.Location = new System.Drawing.Point(3, 512);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(295, 56);
            this.panel1.TabIndex = 6;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button1.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.button1.Location = new System.Drawing.Point(2, 11);
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
            this.button2.Location = new System.Drawing.Point(141, 11);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(147, 35);
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
            this.groupBox3.Size = new System.Drawing.Size(559, 504);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Portfolio list";
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
            this.date,
            this.asset,
            this.types,
            this.Amount_invested,
            this.Share_unit});
            this.dataGridView21.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView21.GridColor = System.Drawing.SystemColors.ButtonShadow;
            this.dataGridView21.Location = new System.Drawing.Point(3, 18);
            this.dataGridView21.Name = "dataGridView21";
            this.dataGridView21.RowHeadersVisible = false;
            this.dataGridView21.RowHeadersWidth = 51;
            this.dataGridView21.RowTemplate.Height = 24;
            this.dataGridView21.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dataGridView21.Size = new System.Drawing.Size(553, 483);
            this.dataGridView21.TabIndex = 0;
          //  this.dataGridView21.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView21_CellContentClick);
            // 
            // date
            // 
            this.date.HeaderText = "Date";
            this.date.MinimumWidth = 6;
            this.date.Name = "date";
            // 
            // asset
            // 
            this.asset.HeaderText = "Asset";
            this.asset.MinimumWidth = 6;
            this.asset.Name = "asset";
            // 
            // types
            // 
            this.types.HeaderText = "Purchase Price";
            this.types.MinimumWidth = 6;
            this.types.Name = "types";
            // 
            // Amount_invested
            // 
            this.Amount_invested.HeaderText = "Amount invested";
            this.Amount_invested.MinimumWidth = 6;
            this.Amount_invested.Name = "Amount_invested";
            // 
            // Share_unit
            // 
            this.Share_unit.HeaderText = "Share Unit";
            this.Share_unit.MinimumWidth = 6;
            this.Share_unit.Name = "Share_unit";
            // 
            // Investment
            // 
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Investment";
            this.Size = new System.Drawing.Size(875, 780);
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
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView21)).EndInit();
            this.ResumeLayout(false);

        }

        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Addinvestment_Click(object sender, EventArgs e)
        {
            string date = Purchase_date.Value.ToString("yyyy-MM-dd");
            string Purchaseasset = Purchase_asset.Text.Trim();
            string Purchaseprice = Purchase_price.Text.Trim();
            string Amountinvest = Amount_invest.Text.Trim();
            string Shareunits = Share_units.Text.Trim();

            // 2️⃣ Validate inputs
            if (string.IsNullOrWhiteSpace(date) || string.IsNullOrWhiteSpace(Purchaseasset) ||
                string.IsNullOrWhiteSpace(Purchaseprice) || string.IsNullOrWhiteSpace(Amountinvest) ||
                string.IsNullOrWhiteSpace(Shareunits))
            {
                MessageBox.Show("Please fill in all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(Amountinvest, out decimal amount))
            {
                MessageBox.Show("Please enter a valid amount.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // 4️⃣ Get next ID
                    int nextId = 1;
                    string idQuery = "SELECT ISNULL(MAX(Id), 0) + 1 FROM Portfolio";

                    using (SqlCommand idCmd = new SqlCommand(idQuery, conn))
                    {
                        nextId = Convert.ToInt32(idCmd.ExecuteScalar());
                    }

                    // 5️⃣ Insert new investment
                    string insertQuery = @"INSERT INTO Portfolio (Id, Asset, Purchase_price, Amount_invested, Share_unit, Date)
                                           VALUES (@Id, @Asset, @Purchase_price, @Amount_invested, @Share_unit, @Date)";

                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", nextId);
                        cmd.Parameters.AddWithValue("@Asset", Purchaseasset);
                        cmd.Parameters.AddWithValue("@Purchase_price", Purchaseprice);
                        cmd.Parameters.AddWithValue("@Amount_invested", Amountinvest);
                        cmd.Parameters.AddWithValue("@Share_unit", Shareunits);
                        cmd.Parameters.AddWithValue("@Date", date);

                        int rows = cmd.ExecuteNonQuery();

                        if (rows > 0)
                        {
                            MessageBox.Show($"Investment added successfully with ID: {nextId}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            // Clear input fields
                            ClearInputFields();
                            LoadportfolioFromDatabase();
                            fractionalshare();
                        }
                        else
                        {
                            MessageBox.Show("Failed to add investment.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // 1️⃣ Check if a row is selected
            if (dataGridView21.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a row to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2️⃣ Confirm deletion
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
                // 3️⃣ Get the selected row
                DataGridViewRow selectedRow = dataGridView21.SelectedRows[0];

                // 4️⃣ Retrieve the Id from the Tag property
                if (selectedRow.Tag == null)
                {
                    MessageBox.Show("Cannot determine the record to delete.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int idToDelete = Convert.ToInt32(selectedRow.Tag);

                // 5️⃣ Delete from database
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string deleteQuery = "DELETE FROM Portfolio WHERE Id = @Id";
                    using (SqlCommand cmd = new SqlCommand(deleteQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", idToDelete);
                        int rows = cmd.ExecuteNonQuery();

                        if (rows > 0)
                        {
                            // 6️⃣ Remove row from DataGridView
                            dataGridView21.Rows.Remove(selectedRow);

                            // Optional: Update total amount or other summaries
                            UpdateTotalAmountInTextBox();
                            fractionalshare();
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

            // Use TRY_CONVERT to safely handle text-based Next_payment values
            string query = @"
        SELECT 
            Date, 
            Asset, 
            Purchase_price, 
            Amount_invested, 
            Share_unit
        FROM Portfolio
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
                                string date = reader["Date"] != DBNull.Value
                                    ? Convert.ToDateTime(reader["Date"]).ToString("yyyy-MM-dd")
                                    : "";
                                string asset = reader["Asset"]?.ToString() ?? "";

                                decimal purchase_price = reader["Purchase_price"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Purchase_price"])
                                    : 0m;
                                decimal amount_invested = reader["Amount_invested"] != DBNull.Value
                                   ? Convert.ToDecimal(reader["Amount_invested"])
                                   : 0m;
                                string share_unit = reader["Share_unit"]?.ToString() ?? "";

                                string formattedAmountinvested = amount_invested.ToString("C");
                                string formattedAmountpurchase = purchase_price.ToString("C");

                                // Add row (without showing Id)
                                int rowIndex = dataGridView21.Rows.Add(date, asset, formattedAmountpurchase, formattedAmountinvested, share_unit);
                            }
                            catch (Exception rowEx)
                            {
                                MessageBox.Show($"Error reading expense record: {rowEx.Message}",
                                                "Row Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
                fractionalshare();
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
            LoadportfolioFromDatabase();
            fractionalshare();
        }
        public void RefreshData()
        {
            LoadportfolioFromDatabase();
            fractionalshare();
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
                saveFileDialog.FileName = "Investment_Report_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";

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

        private void Search_TextChanged(object sender, EventArgs e)
        {
            LoadportfolioFromDatabase(Search.Text.Trim());
        }

        private void LoadportfolioFromDatabase(string searchText = "")
        {
            string query = @"
        SELECT Id, Asset, Purchase_price, Amount_invested, Share_unit, Date
        FROM Portfolio
        WHERE (@search = '' OR Asset LIKE '%' + @search + '%')
        ORDER BY Date DESC;
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
                        dataGridView21.Rows.Clear();

                        while (reader.Read())
                        {
                            try
                            {
                                string date = reader["Date"] != DBNull.Value
                                    ? Convert.ToDateTime(reader["Date"]).ToString("yyyy-MM-dd")
                                    : "";
                                string asset = reader["Asset"]?.ToString() ?? "";

                                decimal purchase_price = reader["Purchase_price"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Purchase_price"])
                                    : 0m;
                                decimal amount_invested = reader["Amount_invested"] != DBNull.Value
                                   ? Convert.ToDecimal(reader["Amount_invested"])
                                   : 0m;
                                string share_unit = reader["Share_unit"]?.ToString() ?? "";

                                string formattedAmountinvested = amount_invested.ToString("C");
                                string formattedAmountpurchase = purchase_price.ToString("C");

                                int rowIndex = dataGridView21.Rows.Add(date, asset, formattedAmountpurchase, formattedAmountinvested, share_unit);
                                dataGridView21.Rows[rowIndex].Tag = reader["Id"];
                            }
                            catch (Exception rowEx)
                            {
                                MessageBox.Show($"Error reading portfolio record: {rowEx.Message}",
                                                "Row Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(
                    $"Database error: {ex.Message}\n\nPlease make sure:\n1. SQL Server is running\n2. Database 'personal_finance' exists\n3. Table 'Portfolio' exists",
                    "Database Connection Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading portfolio data: {ex.Message}",
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        //private void dataGridView21_CellContentClick(object sender, DataGridViewCellEventArgs e)
        //{

        //}
    }
}