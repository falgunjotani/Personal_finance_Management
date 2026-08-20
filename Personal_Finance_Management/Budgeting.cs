using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Personal_Finance_Management
{
    public partial class Budgeting : UserControl
    {
        
        private GroupBox groupBox1;
        private DateTimePicker B_date;
        private Button Exit;
        private Button AddBudget;
        private TextBox B_remaining;
        private TextBox B_spent;
        private TextBox B_category;
        private Label label5;
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
        private DateTimePicker dateTimePicker21;
        private Label label9;
        private GroupBox Total;
        private Label label8;
        private TextBox Income_total;
        private DateTimePicker dateTimePicker11;
        private Panel panel1;
        private Button button1;
        private Button button2;
        private DataGridView dataGridView2;
        
        private System.Windows.Forms.DataVisualization.Charting.Chart pichart;
        private DataGridViewTextBoxColumn SN;
        private DataGridViewTextBoxColumn Typess;
        private DataGridViewTextBoxColumn Amounts;
        private DataGridViewTextBoxColumn date;
        private DataGridViewTextBoxColumn category;
        private DataGridViewTextBoxColumn spent;
        private DataGridViewTextBoxColumn remaining;
      
        string connectionString = AppData.ConnectionString;
        public Budgeting()
        {
            InitializeComponent();
            CreateResponsiveBudgeting();
        }

        private void CreateResponsiveBudgeting()
        {
            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // ✅ Center all text (cells + headers)
            dataGridView2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            LoadBudgetFromDatabase();
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // ✅ Center all text (cells + headers)
            dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            //LoadIncomeTypeFromDatabase();
            //LoadIncomeTypeFromDatabase();
            loaddatacategory();
            DisplayPieChart();
            UpdateTotalAmountInTextBox();
        }

        public void DisplayPieChart()
        {
            // Clear any existing series
            pichart.Series.Clear();

            // Create a new Series for the pie chart
            Series series = new Series
            {
                Name = "PieSeries",
                ChartType = SeriesChartType.Pie,
                IsValueShownAsLabel = true
            };

            // Loop through dataGridView2 rows
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.IsNewRow) continue;

                // Get category name
                string category = row.Cells[1].Value?.ToString() ?? "Unknown";

                // Get spent amount (Cell[2] is the numeric "Spent" value)
                decimal spent = 0m;
                if (row.Cells[2].Value != null)
                {
                    decimal.TryParse(row.Cells[2].Value.ToString(), out spent);
                }

                // Add a data point: Label shows "Category - Amount"
                string label = $"{category} - {spent:C}";
                series.Points.AddXY(label, spent);
            }

            // Add the series to the chart
            pichart.Series.Add(series);

            // Enable legend
            if (pichart.Legends.Count > 0)
                pichart.Legends[0].Enabled = true;
        }



        private void LoadBudgetFromDatabase()
        {
            string query = @"
        SELECT Id, Category, Spent, Remaining, Date
        FROM Budget
        WHERE MONTH(Date) = MONTH(GETDATE())
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
                        dataGridView2.Rows.Clear();

                        while (reader.Read())
                        {
                            try
                            {
                                string date = reader["Date"] != DBNull.Value
                                    ? Convert.ToDateTime(reader["Date"]).ToString("yyyy-MM-dd")
                                    : "";
                                string category = reader["Category"]?.ToString() ?? "";
                                decimal spent = reader["Spent"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Spent"])
                                    : 0m;
                                decimal Remaining = reader["Remaining"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Remaining"])
                                    : 0m;
                                //string frequency = reader["Frequency"]?.ToString() ?? "";
                                 

                                string formattedspent = spent.ToString("C");
                                string formattedRemaining = Remaining.ToString("C");

                                // Add row (without showing Id)
                                int rowIndex = dataGridView2.Rows.Add(date, category, spent, formattedRemaining, formattedRemaining);

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

        private void loaddatacategory()
        {
            try
            {
                // Clear previous rows in dataGridView1 before populating new data
                dataGridView1.Rows.Clear();

                // Iterate through each row of dataGridView2
                int sn = 1;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (row.IsNewRow) continue; // Skip the new row placeholder

                    // Get Type from Cell[1]
                    string categ = row.Cells[1].Value?.ToString() ?? "";

                    // Get the numeric values from Cell[2] and Cell[3]
                    string cell2 = row.Cells[2].Value?.ToString() ?? "0";
                    string cell3 = row.Cells[3].Value?.ToString() ?? "0";

                    // Remove non-numeric characters (like ¥, commas, spaces)
                    string cleanCell2 = new string(cell2.Where(c => char.IsDigit(c) || c == '.' || c == '-').ToArray());
                    string cleanCell3 = new string(cell3.Where(c => char.IsDigit(c) || c == '.' || c == '-').ToArray());

                    // Parse to decimal
                    if (decimal.TryParse(cleanCell2, out decimal val2) && decimal.TryParse(cleanCell3, out decimal val3))
                    {
                        decimal difference = val2 - val3;

                        // Add row to dataGridView1
                        dataGridView1.Rows.Add(sn, categ, difference);
                        sn++;
                    }
                    else
                    {
                        // If parsing fails, log and show zero
                        Console.WriteLine($"Invalid numeric values at row {sn}: '{cell2}' and '{cell3}'");
                        dataGridView1.Rows.Add(sn, categ, 0);
                        sn++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading expense data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
        private void InitializeComponent()
        {
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea3 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend3 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.B_date = new System.Windows.Forms.DateTimePicker();
            this.Exit = new System.Windows.Forms.Button();
            this.AddBudget = new System.Windows.Forms.Button();
            this.B_remaining = new System.Windows.Forms.TextBox();
            this.B_spent = new System.Windows.Forms.TextBox();
            this.B_category = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.category = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.spent = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.remaining = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pichart = new System.Windows.Forms.DataVisualization.Charting.Chart();
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
            this.dateTimePicker21 = new System.Windows.Forms.DateTimePicker();
            this.label9 = new System.Windows.Forms.Label();
            this.Total = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.Income_total = new System.Windows.Forms.TextBox();
            this.dateTimePicker11 = new System.Windows.Forms.DateTimePicker();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pichart)).BeginInit();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.Reports.SuspendLayout();
            this.Total.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.B_date);
            this.groupBox1.Controls.Add(this.Exit);
            this.groupBox1.Controls.Add(this.AddBudget);
            this.groupBox1.Controls.Add(this.B_remaining);
            this.groupBox1.Controls.Add(this.B_spent);
            this.groupBox1.Controls.Add(this.B_category);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(0, 586);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(870, 188);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Add  BudgetSchedule";
            // 
            // B_date
            // 
            this.B_date.Font = new System.Drawing.Font("Segoe UI", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.B_date.Location = new System.Drawing.Point(232, 30);
            this.B_date.Name = "B_date";
            this.B_date.Size = new System.Drawing.Size(249, 31);
            this.B_date.TabIndex = 0;
            this.B_date.Value = new System.DateTime(2025, 11, 2, 14, 54, 39, 0);
            // 
            // Exit
            // 
            this.Exit.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.Exit.BackColor = System.Drawing.Color.OrangeRed;
            this.Exit.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.Exit.Location = new System.Drawing.Point(711, 117);
            this.Exit.Name = "Exit";
            this.Exit.Size = new System.Drawing.Size(134, 55);
            this.Exit.TabIndex = 11;
            this.Exit.Text = "Exit";
            this.Exit.UseVisualStyleBackColor = false;
            this.Exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // AddBudget
            // 
            this.AddBudget.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.AddBudget.BackColor = System.Drawing.Color.LimeGreen;
            this.AddBudget.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.AddBudget.Location = new System.Drawing.Point(540, 117);
            this.AddBudget.Name = "AddBudget";
            this.AddBudget.Size = new System.Drawing.Size(143, 55);
            this.AddBudget.TabIndex = 10;
            this.AddBudget.Text = "Add ";
            this.AddBudget.UseVisualStyleBackColor = false;
            this.AddBudget.Click += new System.EventHandler(this.Addexpenses_Click);
            // 
            // B_remaining
            // 
            this.B_remaining.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.B_remaining.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.B_remaining.Location = new System.Drawing.Point(232, 122);
            this.B_remaining.Name = "B_remaining";
            this.B_remaining.Size = new System.Drawing.Size(249, 30);
            this.B_remaining.TabIndex = 7;
            // 
            // B_spent
            // 
            this.B_spent.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.B_spent.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.B_spent.Location = new System.Drawing.Point(232, 92);
            this.B_spent.Name = "B_spent";
            this.B_spent.Size = new System.Drawing.Size(249, 30);
            this.B_spent.TabIndex = 6;
            // 
            // B_category
            // 
            this.B_category.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.B_category.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.B_category.Location = new System.Drawing.Point(232, 62);
            this.B_category.Name = "B_category";
            this.B_category.Size = new System.Drawing.Size(249, 30);
            this.B_category.TabIndex = 5;
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
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(56, 119);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(67, 28);
            this.label3.TabIndex = 2;
            this.label3.Text = "Spent:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(56, 91);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(87, 28);
            this.label2.TabIndex = 1;
            this.label2.Text = "Amount:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(56, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 28);
            this.label1.TabIndex = 0;
            this.label1.Text = "Category:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dataGridView2);
            this.groupBox2.Controls.Add(this.pichart);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.groupBox4);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(870, 586);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Budget schedule Details";
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
            this.date,
            this.category,
            this.spent,
            this.remaining});
            this.dataGridView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView2.GridColor = System.Drawing.SystemColors.ButtonShadow;
            this.dataGridView2.Location = new System.Drawing.Point(3, 18);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowHeadersVisible = false;
            this.dataGridView2.RowHeadersWidth = 51;
            this.dataGridView2.RowTemplate.Height = 24;
            this.dataGridView2.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.dataGridView2.Size = new System.Drawing.Size(537, 323);
            this.dataGridView2.TabIndex = 0;
            // 
            // date
            // 
            this.date.HeaderText = "Date";
            this.date.MinimumWidth = 6;
            this.date.Name = "date";
            // 
            // category
            // 
            this.category.HeaderText = "Category";
            this.category.MinimumWidth = 6;
            this.category.Name = "category";
            // 
            // spent
            // 
            this.spent.HeaderText = "Amount ";
            this.spent.MinimumWidth = 6;
            this.spent.Name = "spent";
            // 
            // remaining
            // 
            this.remaining.HeaderText = "Spent";
            this.remaining.MinimumWidth = 6;
            this.remaining.Name = "remaining";
            // 
            // pichart
            // 
            chartArea3.Name = "ChartArea1";
            this.pichart.ChartAreas.Add(chartArea3);
            this.pichart.Dock = System.Windows.Forms.DockStyle.Bottom;
            legend3.Name = "Legend1";
            this.pichart.Legends.Add(legend3);
            this.pichart.Location = new System.Drawing.Point(3, 341);
            this.pichart.Name = "pichart";
            this.pichart.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.None;
            series3.ChartArea = "ChartArea1";
            series3.Legend = "Legend1";
            series3.Name = "Series1";
            this.pichart.Series.Add(series3);
            this.pichart.Size = new System.Drawing.Size(537, 242);
            this.pichart.TabIndex = 13;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(57, 29);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(254, 41);
            this.label6.TabIndex = 2;
            this.label6.Text = "Budget Planning";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.dataGridView1);
            this.groupBox4.Controls.Add(this.label7);
            this.groupBox4.Controls.Add(this.Reports);
            this.groupBox4.Controls.Add(this.panel1);
            this.groupBox4.Dock = System.Windows.Forms.DockStyle.Right;
            this.groupBox4.Location = new System.Drawing.Point(540, 18);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(327, 565);
            this.groupBox4.TabIndex = 1;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Budget Analysis";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeColumns = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
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
            this.Typess.HeaderText = "Category";
            this.Typess.MinimumWidth = 6;
            this.Typess.Name = "Typess";
            // 
            // Amounts
            // 
            this.Amounts.FillWeight = 117.9144F;
            this.Amounts.HeaderText = "Remaining";
            this.Amounts.MinimumWidth = 6;
            this.Amounts.Name = "Amounts";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold);
            this.label7.ForeColor = System.Drawing.Color.Firebrick;
            this.label7.Location = new System.Drawing.Point(105, 18);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(166, 23);
            this.label7.TabIndex = 11;
            this.label7.Text = "Remaining Amount";
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
            this.Reports.Location = new System.Drawing.Point(3, 308);
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
            this.label8.Location = new System.Drawing.Point(15, 26);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(101, 23);
            this.label8.TabIndex = 9;
            this.label8.Text = "Total Spent";
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
            this.panel1.Location = new System.Drawing.Point(3, 506);
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
            // Budgeting
            // 
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Budgeting";
            this.Size = new System.Drawing.Size(870, 774);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pichart)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.Reports.ResumeLayout(false);
            this.Reports.PerformLayout();
            this.Total.ResumeLayout(false);
            this.Total.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        

        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void Addexpenses_Click(object sender, EventArgs e)
        {
            string date = B_date.Value.ToString("yyyy-MM-dd");
            string category = B_category.Text.Trim();
            string spentText = B_spent.Text.Trim();
            string remainingText = B_remaining.Text.Trim();

            // 1️⃣ Validate inputs
            if (string.IsNullOrWhiteSpace(category) ||
                string.IsNullOrWhiteSpace(spentText) ||
                string.IsNullOrWhiteSpace(remainingText))
            {
                MessageBox.Show("Please fill in all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(spentText, out decimal spent))
            {
                MessageBox.Show("Please enter a valid amount for Spent.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(remainingText, out decimal remaining))
            {
                MessageBox.Show("Please enter a valid amount for Remaining.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // 2️⃣ Check if category already exists
                    string checkQuery = "SELECT Id, Spent, Remaining FROM Budget WHERE Category = @Category";
                    using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                    {
                        checkCmd.Parameters.AddWithValue("@Category", category);
                        using (SqlDataReader reader = checkCmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // Category exists → update Spent and optionally Remaining
                                int existingId = Convert.ToInt32(reader["Id"]);
                                decimal existingSpent = reader["Spent"] != DBNull.Value ? Convert.ToDecimal(reader["Spent"]) : 0m;
                                decimal existingRemaining = reader["Remaining"] != DBNull.Value ? Convert.ToDecimal(reader["Remaining"]) : remaining;

                                reader.Close();

                                // Update Spent and Remaining
                                string updateQuery = "UPDATE Budget SET Spent = @NewSpent, Remaining = @NewRemaining, Date = @Date WHERE Id = @Id";
                                using (SqlCommand updateCmd = new SqlCommand(updateQuery, conn))
                                {
                                    updateCmd.Parameters.AddWithValue("@NewSpent", existingSpent + spent);
                                    updateCmd.Parameters.AddWithValue("@NewRemaining", existingRemaining+ remaining); // optionally: existingRemaining - spent
                                    updateCmd.Parameters.AddWithValue("@Date", date);
                                    updateCmd.Parameters.AddWithValue("@Id", existingId);

                                    int rowsUpdated = updateCmd.ExecuteNonQuery();
                                    if (rowsUpdated > 0)
                                        MessageBox.Show($"Updated existing category '{category}' with new spent amount.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                            else
                            {
                                reader.Close();

                                // Get the next available ID
                                int nextId = 1;
                                string idQuery = "SELECT ISNULL(MAX(Id), 0) + 1 FROM Budget";
                                using (SqlCommand idCmd = new SqlCommand(idQuery, conn))
                                {
                                    nextId = Convert.ToInt32(idCmd.ExecuteScalar());
                                }

                                // Category does not exist → insert new row with the generated ID
                                string insertQuery = @"INSERT INTO Budget (Id, Category, Spent, Remaining, Date) 
                                       VALUES (@Id, @Category, @Spent, @Remaining, @Date)";
                                using (SqlCommand insertCmd = new SqlCommand(insertQuery, conn))
                                {
                                    insertCmd.Parameters.AddWithValue("@Id", nextId);
                                    insertCmd.Parameters.AddWithValue("@Category", category);
                                    insertCmd.Parameters.AddWithValue("@Spent", spent);
                                    insertCmd.Parameters.AddWithValue("@Remaining", remaining);
                                    insertCmd.Parameters.AddWithValue("@Date", date);

                                    int rowsInserted = insertCmd.ExecuteNonQuery();
                                    if (rowsInserted > 0)
                                        MessageBox.Show($"Added new category '{category}' successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                        }
                    }

                    // 3️⃣ Clear inputs and refresh UI
                    B_category.Clear();
                    B_spent.Clear();
                    B_remaining.Clear();
                    LoadBudgetFromDatabase();
                    loaddatacategory();
                    DisplayPieChart();
                    UpdateTotalAmountInTextBox();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

            // Use TRY_CONVERT to safely handle text-based Date values
            string query = @"
        SELECT 
            Id,
            Category, 
            Spent, 
            Remaining, 
            Date           
        FROM Budget
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
                        dataGridView2.Rows.Clear();

                        while (reader.Read())
                        {
                            try
                            {
                                string date = reader["Date"] != DBNull.Value
                                    ? Convert.ToDateTime(reader["Date"]).ToString("yyyy-MM-dd")
                                    : "";
                                string category = reader["Category"]?.ToString() ?? "";
                                decimal spent = reader["Spent"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Spent"])
                                    : 0m;
                                decimal remaining = reader["Remaining"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Remaining"])
                                    : 0m;

                                string formattedSpent = spent.ToString("C");
                                string formattedRemaining = remaining.ToString("C");

                                // Add row (without showing Id) - match the same column structure as LoadBudgetFromDatabase
                                int rowIndex = dataGridView2.Rows.Add(date, category, spent, formattedRemaining, formattedRemaining);

                                // Store the Id in the Tag property
                                dataGridView2.Rows[rowIndex].Tag = reader["Id"];
                            }
                            catch (Exception rowEx)
                            {
                                MessageBox.Show($"Error reading budget record: {rowEx.Message}",
                                                "Row Error",
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Warning);
                            }
                        }
                    }
                }

                // Optional: Refresh other UI components if needed
                // LoadBudgetFromDatabase(); // Remove this line as it will overwrite your filtered data
                // loaddatacategory();
                // DisplayPieChart();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(
                    $"Database error: {ex.Message}\n\nPlease make sure:\n1. SQL Server is running\n2. Database exists\n3. Table 'Budget' exists",
                    "Database Connection Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading budget data: {ex.Message}",
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        private void View_all_Click(object sender, EventArgs e)
        {
           // CreateResponsiveBudgeting();
            LoadBudgetFromDatabase();
            loaddatacategory();
            DisplayPieChart();
            UpdateTotalAmountInTextBox();

        }
        public void RefreshData()
        {
            // Reload data from database
            CreateResponsiveBudgeting();
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
                saveFileDialog.FileName = "Budget_Report_" + DateTime.Now.ToString("yyyyMMdd") + ".csv";

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

                    string deleteQuery = "DELETE FROM Budget WHERE Id = @Id";
                    using (SqlCommand cmd = new SqlCommand(deleteQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", idToDelete);
                        int rows = cmd.ExecuteNonQuery();

                        if (rows > 0)
                        {
                            // Remove row from DataGridView
                            dataGridView2.Rows.Remove(selectedRow);

                            LoadBudgetFromDatabase();
                            loaddatacategory();
                            DisplayPieChart();
                            UpdateTotalAmountInTextBox();

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
        //private void Budgeting_Load(object sender, EventArgs e)
        //{

        //}
    }
}