using System;
using System.Collections.Generic;
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

    public partial class Saving : UserControl
    {
       
        private GroupBox groupBox1;
        private DateTimePicker Deadline_saving;
        private Button Exit;
        private Button Addsaving;
        private TextBox Saved_saving;
        private TextBox Target_saving;
        private TextBox Goal_saving;
        private Label label5;
        private Label label3;
        private Label label2;
        private Label label1;
        private GroupBox groupBox2;
        private Label label6;
        private GroupBox groupBox3;
        private DataGridView dataGridView21;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private GroupBox search;
        private DataGridView dataGridView11;
        private DataGridViewTextBoxColumn SN;
        private DataGridViewTextBoxColumn Goals;
        private DataGridViewTextBoxColumn pprogress;
        private Label label7;
        private Panel panel1;
        private Button button1;
        private Button button2;
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
        private DataGridViewTextBoxColumn Date;
        private DataGridViewTextBoxColumn Goal;
        private DataGridViewTextBoxColumn Target;
        private DataGridViewTextBoxColumn Saved;
        private DataGridViewTextBoxColumn Progress;
        private Label label4;
        private TextBox Tsaving;
        string connectionString = AppData.ConnectionString;
      
        public Saving()
        {
            InitializeComponent();
            CreateResponsiveSaving();
        }

        private void CreateResponsiveSaving()
        {
            dataGridView21.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView21.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
         dataGridView21.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
         dataGridView21.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
         dataGridView21.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
         dataGridView21.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            LoadsavingFromDatabase();
            dataGridView11.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView11.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // ✅ Center all text (cells + headers)
            dataGridView11.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView11.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView11.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            dataGridView11.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            LoadexpensessavingFromDatabase();
            
            //chart1.Titles.Add("Quarterly Sales Overview");
        }

        private void LoadsavingFromDatabase()
        {
            string query = "SELECT Id, Goal, Target, Saved, Progress, Deadline FROM Saving ORDER BY Deadline DESC;";

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
                                string goal = reader["Goal"]?.ToString() ?? "";
                                string target = reader["Target"]?.ToString() ?? "";
                                decimal saved = reader["Saved"] != DBNull.Value
                                   ? Convert.ToDecimal(reader["Saved"])
                                   : 0m;
                                string progress = reader["Progress"]?.ToString() ?? "";
                                string deadline = reader["Deadline"] != DBNull.Value
                                    ? Convert.ToDateTime(reader["Deadline"]).ToString("yyyy-MM-dd")
                                    : "";
                                
                                
                                
                               





                                string formattedAmount = saved.ToString("C");

                                // Add row (without showing Id)
                                int rowIndex = dataGridView21.Rows.Add(deadline,goal, target, formattedAmount, progress);

                                // Store the Id in the Tag property
                                dataGridView21.Rows[rowIndex].Tag = reader["Id"];
                            }
                            catch (Exception rowEx)
                            {
                                MessageBox.Show($"Error reading income record: {rowEx.Message}",
                                                "Row Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                    CreateGoalProgressChart();
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

        private void CreateGoalProgressChart()
        {
            // 1. Clean start
            chart1.Series.Clear();
            chart1.ChartAreas.Clear();
            chart1.Titles.Clear();

            // 2. ChartArea – fill the control
            ChartArea area = new ChartArea("MainArea");
            chart1.ChartAreas.Add(area);

            area.Position = new ElementPosition(5, 5, 90, 85);
            area.InnerPlotPosition = new ElementPosition(12, 10, 80, 80);

            // 3. Axes
            area.AxisX.Title = "Goals";
            area.AxisX.Interval = 1;
            area.AxisX.MajorGrid.Enabled = false;
            //area.AxisX.LabelStyle.Angle = -45; // Uncomment if you want angled labels

            area.AxisY.Title = "Progress (%)";
            area.AxisY.Minimum = 0;
            area.AxisY.Maximum = 100;
            area.AxisY.Interval = 20;
            area.AxisY.MajorGrid.Enabled = true;
            area.AxisY.MajorGrid.LineColor = Color.LightGray;

            // 4. Series – inline, colour per point
            chart1.Series.Add(new Series
            {
                Name = "GoalProgress",
                ChartType = SeriesChartType.Column,
                IsValueShownAsLabel = true,
                LabelFormat = "{0}%",
                BorderWidth = 1,
                BorderColor = Color.DarkBlue,
                ["PointWidth"] = "0.8"
            });

            // 5. Data points - Get data from DataGridView
            var points = chart1.Series["GoalProgress"].Points;

            foreach (DataGridViewRow row in dataGridView21.Rows)
            {
                // Skip empty rows
                if (row.IsNewRow) continue;

                try
                {
                    // Get goal name from first column (index 0)
                    string goal = row.Cells[1].Value?.ToString() ?? "Unknown Goal";

                    // Get progress from fourth column (index 3) and convert to double
                    string progressText = row.Cells[4].Value?.ToString() ?? "0";

                    // Remove percentage sign if present and convert to double
                    if (progressText.Contains("%"))
                    {
                        progressText = progressText.Replace("%", "");
                    }

                    if (double.TryParse(progressText, out double progressValue))
                    {
                        points.AddXY(goal, progressValue);
                    }
                    else
                    {
                        // If parsing fails, add with 0 progress
                        points.AddXY(goal, 0);
                    }
                }
                catch (Exception ex)
                {
                    // Log error or handle silently
                    System.Diagnostics.Debug.WriteLine($"Error adding chart point: {ex.Message}");
                }
            }

            // 6. Colour rule
            foreach (DataPoint pt in points)
            {
                double v = pt.YValues[0];
                pt.Color = v > 90 ? Color.LimeGreen :
                           v < 50 ? Color.IndianRed :
                                    Color.Gold;
            }

            // 7. Title
            chart1.Titles.Add(new Title(
                "Financial Goal Progress",
                Docking.Top,
                new Font("Arial", 14, FontStyle.Bold),
                Color.Black));
        }
        private void LoadexpensessavingFromDatabase()
        {
            try
            {
                // Clear previous rows in dataGridView11 before populating new data
                dataGridView11.Rows.Clear();

                // Iterate through each row of dataGridView21
                int sn = 1;
                foreach (DataGridViewRow row in dataGridView21.Rows)
                {
                    if (row.IsNewRow) continue; // Skip the new row placeholder

                    // Get Type from Cell[1] and Progress from Cell[4]
                    string categ = row.Cells[1].Value?.ToString() ?? ""; // Assuming "Type" is at index 1
                    string progress = row.Cells[4].Value?.ToString() ?? ""; // Assuming "Progress" is at index 4

                    // Log the progress value for debugging
                    Console.WriteLine($"Row {sn} - Progress: {progress}");

                    // Remove any non-numeric characters like '%' from the progress string (if present)
                    progress = progress.Replace("%", "").Trim();

                    // Try parsing the progress value to a decimal
                    if (decimal.TryParse(progress, out decimal progressValue))
                    {
                        // Only add to dataGridView11 if progress is less than 50%
                        if (progressValue < 50)
                        {
                            dataGridView11.Rows.Add(sn, categ, progress);
                            sn++;
                        }
                    }
                    else
                    {
                        // Log error if progress cannot be parsed to a decimal
                        Console.WriteLine($"Invalid progress value: {progress} at row {sn}");
                    }
                }
                UpdateTotalAmountInTextBox();
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
                decimal totalAmount1 = 0m;

                foreach (DataGridViewRow row in dataGridView21.Rows)
                {
                    if (row.IsNewRow) continue; // Skip the new row placeholder

                    string amountStr = row.Cells[3].Value?.ToString() ?? "$0"; // Assuming "Amount" is at index 3
                    string amountStr1 = row.Cells[2].Value?.ToString() ?? "$0"; // Assuming "Amount" is at index 3

                    // Remove currency symbol and parse
                    if (decimal.TryParse(amountStr, System.Globalization.NumberStyles.Currency, null, out decimal amount))
                    {
                        totalAmount += amount;
                    }
                    if (decimal.TryParse(amountStr1, System.Globalization.NumberStyles.Currency, null, out decimal amount1))
                    {
                        totalAmount1 += amount1;
                    }
                }

                // Display the total in your TextBox
                Income_total.Text = totalAmount.ToString("C"); // "C" formats as currency
                Tsaving.Text = totalAmount1.ToString("C"); // "C" formats as currency
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error calculating total amount: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

      
         

        private Panel CreateGoalProgressItem(string goal, int progress, Color color)
        {
            Panel panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.Margin = new Padding(0, 8, 0, 8);
            panel.Height = 60;

            Label goalLabel = new Label();
            goalLabel.Text = goal;
            goalLabel.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            goalLabel.Location = new Point(0, 5);
            goalLabel.AutoSize = true;

            Panel bgPanel = new Panel();
            bgPanel.BackColor = Color.LightGray;
            bgPanel.Location = new Point(0, 30);
            bgPanel.Size = new Size(200, 20);

            Panel progressPanel = new Panel();
            progressPanel.BackColor = color;
            progressPanel.Location = new Point(0, 0);
            progressPanel.Size = new Size((int)(200 * progress / 100.0), 20);
            bgPanel.Controls.Add(progressPanel);

            Label percentLabel = new Label();
            percentLabel.Text = $"{progress}%";
            percentLabel.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            percentLabel.ForeColor = color;
            percentLabel.Location = new Point(210, 30);
            percentLabel.AutoSize = true;

            panel.Controls.Add(goalLabel);
            panel.Controls.Add(bgPanel);
            panel.Controls.Add(percentLabel);

            return panel;
        }

        
        private void InitializeComponent()
        {
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.Deadline_saving = new System.Windows.Forms.DateTimePicker();
            this.Exit = new System.Windows.Forms.Button();
            this.Addsaving = new System.Windows.Forms.Button();
            this.Saved_saving = new System.Windows.Forms.TextBox();
            this.Target_saving = new System.Windows.Forms.TextBox();
            this.Goal_saving = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.label6 = new System.Windows.Forms.Label();
            this.search = new System.Windows.Forms.GroupBox();
            this.dataGridView11 = new System.Windows.Forms.DataGridView();
            this.SN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Goals = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pprogress = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label7 = new System.Windows.Forms.Label();
            this.Reports = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.View_all = new System.Windows.Forms.Button();
            this.View_reports = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.dateTimePicker21 = new System.Windows.Forms.DateTimePicker();
            this.label9 = new System.Windows.Forms.Label();
            this.Total = new System.Windows.Forms.GroupBox();
            this.Tsaving = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.Income_total = new System.Windows.Forms.TextBox();
            this.dateTimePicker11 = new System.Windows.Forms.DateTimePicker();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dataGridView21 = new System.Windows.Forms.DataGridView();
            this.Date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Goal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Target = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Saved = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Progress = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.search.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView11)).BeginInit();
            this.Reports.SuspendLayout();
            this.panel1.SuspendLayout();
            this.Total.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView21)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.Deadline_saving);
            this.groupBox1.Controls.Add(this.Exit);
            this.groupBox1.Controls.Add(this.Addsaving);
            this.groupBox1.Controls.Add(this.Saved_saving);
            this.groupBox1.Controls.Add(this.Target_saving);
            this.groupBox1.Controls.Add(this.Goal_saving);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(0, 807);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1094, 167);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Add  Saving";
            // 
            // Deadline_saving
            // 
            this.Deadline_saving.Font = new System.Drawing.Font("Segoe UI", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Deadline_saving.Location = new System.Drawing.Point(193, 30);
            this.Deadline_saving.Name = "Deadline_saving";
            this.Deadline_saving.Size = new System.Drawing.Size(254, 31);
            this.Deadline_saving.TabIndex = 0;
            this.Deadline_saving.Value = new System.DateTime(2025, 11, 2, 14, 54, 39, 0);
            // 
            // Exit
            // 
            this.Exit.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.Exit.BackColor = System.Drawing.Color.OrangeRed;
            this.Exit.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.Exit.Location = new System.Drawing.Point(941, 97);
            this.Exit.Name = "Exit";
            this.Exit.Size = new System.Drawing.Size(134, 55);
            this.Exit.TabIndex = 11;
            this.Exit.Text = "Exit";
            this.Exit.UseVisualStyleBackColor = false;
            this.Exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // Addsaving
            // 
            this.Addsaving.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.Addsaving.BackColor = System.Drawing.Color.LimeGreen;
            this.Addsaving.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.Addsaving.Location = new System.Drawing.Point(780, 97);
            this.Addsaving.Name = "Addsaving";
            this.Addsaving.Size = new System.Drawing.Size(143, 55);
            this.Addsaving.TabIndex = 10;
            this.Addsaving.Text = "Add Saving";
            this.Addsaving.UseVisualStyleBackColor = false;
            this.Addsaving.Click += new System.EventHandler(this.Addsaving_Click);
            // 
            // Saved_saving
            // 
            this.Saved_saving.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Saved_saving.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Saved_saving.Location = new System.Drawing.Point(193, 122);
            this.Saved_saving.Name = "Saved_saving";
            this.Saved_saving.Size = new System.Drawing.Size(473, 30);
            this.Saved_saving.TabIndex = 7;
            // 
            // Target_saving
            // 
            this.Target_saving.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Target_saving.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Target_saving.Location = new System.Drawing.Point(193, 94);
            this.Target_saving.Name = "Target_saving";
            this.Target_saving.Size = new System.Drawing.Size(473, 30);
            this.Target_saving.TabIndex = 6;
            // 
            // Goal_saving
            // 
            this.Goal_saving.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.Goal_saving.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Goal_saving.Location = new System.Drawing.Point(193, 64);
            this.Goal_saving.Name = "Goal_saving";
            this.Goal_saving.Size = new System.Drawing.Size(473, 30);
            this.Goal_saving.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(59, 33);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(89, 28);
            this.label5.TabIndex = 4;
            this.label5.Text = "Deadline";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(56, 119);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(69, 28);
            this.label3.TabIndex = 2;
            this.label3.Text = "Saved:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(56, 91);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 28);
            this.label2.TabIndex = 1;
            this.label2.Text = "Target:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(56, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 28);
            this.label1.TabIndex = 0;
            this.label1.Text = "Goal:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chart1);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.search);
            this.groupBox2.Controls.Add(this.groupBox3);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(1094, 807);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Saving Details";
            // 
            // chart1
            // 
            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            this.chart1.Dock = System.Windows.Forms.DockStyle.Bottom;
            legend1.Name = "Legend1";
            this.chart1.Legends.Add(legend1);
            this.chart1.Location = new System.Drawing.Point(3, 513);
            this.chart1.Name = "chart1";
            series1.ChartArea = "ChartArea1";
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            this.chart1.Series.Add(series1);
            this.chart1.Size = new System.Drawing.Size(761, 291);
            this.chart1.TabIndex = 13;
            this.chart1.Text = "chart1";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(54, 21);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(218, 41);
            this.label6.TabIndex = 2;
            this.label6.Text = "Saving Details";
            // 
            // search
            // 
            this.search.Controls.Add(this.dataGridView11);
            this.search.Controls.Add(this.label7);
            this.search.Controls.Add(this.Reports);
            this.search.Dock = System.Windows.Forms.DockStyle.Right;
            this.search.Location = new System.Drawing.Point(764, 18);
            this.search.Name = "search";
            this.search.Size = new System.Drawing.Size(327, 786);
            this.search.TabIndex = 1;
            this.search.TabStop = false;
            this.search.Text = "Saving Categories";
            // 
            // dataGridView11
            // 
            this.dataGridView11.AllowUserToAddRows = false;
            this.dataGridView11.AllowUserToDeleteRows = false;
            this.dataGridView11.AllowUserToResizeColumns = false;
            this.dataGridView11.AllowUserToResizeRows = false;
            this.dataGridView11.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView11.BackgroundColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView11.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView11.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView11.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.SN,
            this.Goals,
            this.pprogress});
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
            // Goals
            // 
            this.Goals.FillWeight = 117.9144F;
            this.Goals.HeaderText = "Goal";
            this.Goals.MinimumWidth = 6;
            this.Goals.Name = "Goals";
            // 
            // pprogress
            // 
            this.pprogress.FillWeight = 117.9144F;
            this.pprogress.HeaderText = "Progress";
            this.pprogress.MinimumWidth = 6;
            this.pprogress.Name = "pprogress";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold);
            this.label7.ForeColor = System.Drawing.Color.Brown;
            this.label7.Location = new System.Drawing.Point(21, 18);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(267, 23);
            this.label7.TabIndex = 11;
            this.label7.Text = "Need Improvement(below 50%)";
            // 
            // Reports
            // 
            this.Reports.Controls.Add(this.label4);
            this.Reports.Controls.Add(this.View_all);
            this.Reports.Controls.Add(this.View_reports);
            this.Reports.Controls.Add(this.panel1);
            this.Reports.Controls.Add(this.label10);
            this.Reports.Controls.Add(this.dateTimePicker21);
            this.Reports.Controls.Add(this.label9);
            this.Reports.Controls.Add(this.Total);
            this.Reports.Controls.Add(this.dateTimePicker11);
            this.Reports.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.Reports.Location = new System.Drawing.Point(3, 481);
            this.Reports.Name = "Reports";
            this.Reports.Size = new System.Drawing.Size(321, 302);
            this.Reports.TabIndex = 10;
            this.Reports.TabStop = false;
            this.Reports.Text = "Report";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(9, 187);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(67, 23);
            this.label4.TabIndex = 11;
            this.label4.Text = "Target ";
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
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(3, 243);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(315, 56);
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
            this.Total.Controls.Add(this.Tsaving);
            this.Total.Controls.Add(this.label8);
            this.Total.Controls.Add(this.Income_total);
            this.Total.Location = new System.Drawing.Point(3, 120);
            this.Total.Name = "Total";
            this.Total.Size = new System.Drawing.Size(312, 117);
            this.Total.TabIndex = 8;
            this.Total.TabStop = false;
            this.Total.Text = "Total";
            // 
            // Tsaving
            // 
            this.Tsaving.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.Tsaving.Location = new System.Drawing.Point(130, 64);
            this.Tsaving.Name = "Tsaving";
            this.Tsaving.Size = new System.Drawing.Size(166, 30);
            this.Tsaving.TabIndex = 10;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(6, 26);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(108, 23);
            this.label8.TabIndex = 9;
            this.label8.Text = "Total Saving";
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
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.Controls.Add(this.dataGridView21);
            this.groupBox3.Location = new System.Drawing.Point(6, 55);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(752, 422);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Saving list";
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
            this.Goal,
            this.Target,
            this.Saved,
            this.Progress});
            this.dataGridView21.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView21.GridColor = System.Drawing.SystemColors.ButtonShadow;
            this.dataGridView21.Location = new System.Drawing.Point(3, 18);
            this.dataGridView21.Name = "dataGridView21";
            this.dataGridView21.RowHeadersVisible = false;
            this.dataGridView21.RowHeadersWidth = 51;
            this.dataGridView21.RowTemplate.Height = 24;
            this.dataGridView21.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.dataGridView21.Size = new System.Drawing.Size(746, 401);
            this.dataGridView21.TabIndex = 0;
            // 
            // Date
            // 
            this.Date.HeaderText = "Date";
            this.Date.MinimumWidth = 6;
            this.Date.Name = "Date";
            // 
            // Goal
            // 
            this.Goal.HeaderText = "Goal";
            this.Goal.MinimumWidth = 6;
            this.Goal.Name = "Goal";
            // 
            // Target
            // 
            this.Target.HeaderText = "Target";
            this.Target.MinimumWidth = 6;
            this.Target.Name = "Target";
            // 
            // Saved
            // 
            this.Saved.HeaderText = "Saved";
            this.Saved.MinimumWidth = 6;
            this.Saved.Name = "Saved";
            // 
            // Progress
            // 
            this.Progress.HeaderText = "Progress";
            this.Progress.MinimumWidth = 6;
            this.Progress.Name = "Progress";
            // 
            // Saving
            // 
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Saving";
            this.Size = new System.Drawing.Size(1094, 974);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.search.ResumeLayout(false);
            this.search.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView11)).EndInit();
            this.Reports.ResumeLayout(false);
            this.Reports.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.Total.ResumeLayout(false);
            this.Total.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView21)).EndInit();
            this.ResumeLayout(false);

        }

        
        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void View_all_Click(object sender, EventArgs e)
        {
            LoadsavingFromDatabase(); LoadexpensessavingFromDatabase(); CreateGoalProgressChart();
        }
        public void RefreshData()
        {
            LoadsavingFromDatabase(); LoadexpensessavingFromDatabase(); CreateGoalProgressChart();
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

            // Use TRY_CONVERT to safely handle text-based Next_payment values
            string query = @"
        SELECT 
            Deadline, 
            Goal, 
            Target, 
            Saved, 
            Progress
        FROM Saving
        WHERE TRY_CONVERT(date, Deadline) BETWEEN @FromDate AND @ToDate
        ORDER BY TRY_CONVERT(date, Deadline) DESC;
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

                                string description = reader["Goal"]?.ToString() ?? "";
                                //string category = reader["Category"]?.ToString() ?? "";
                                decimal target = reader["Target"] != DBNull.Value
                                    ? Convert.ToDecimal(reader["Target"])
                                    : 0m;
                                decimal saved = reader["Saved"] != DBNull.Value
                                   ? Convert.ToDecimal(reader["Saved"])
                                   : 0m;
                                string progress = reader["Progress"]?.ToString() ?? "";
                                // Safely parse Next_payment
                                string nextPaymentStr = reader["Deadline"]?.ToString();
                                string formattedNextPayment = "";

                                if (DateTime.TryParse(nextPaymentStr, out DateTime nextPaymentDate))
                                    formattedNextPayment = nextPaymentDate.ToString("yyyy-MM-dd");

                                string formattedtargett = target.ToString("C");
                                string formattedsaved = saved.ToString("C");

                                // Add to DataGridView
                                dataGridView21.Rows.Add(formattedNextPayment, description, formattedtargett, formattedsaved, progress);
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
                LoadexpensessavingFromDatabase();
                CreateGoalProgressChart();
                // LoadexpensescategoryFromDatabase();
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

                    string deleteQuery = "DELETE FROM Saving WHERE Id = @Id";
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
                            LoadsavingFromDatabase(); LoadexpensessavingFromDatabase(); CreateGoalProgressChart();

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

        private void Addsaving_Click(object sender, EventArgs e)
        {
            string date = Deadline_saving.Value.ToString("yyyy-MM-dd");
            string goal = Goal_saving.Text.Trim();
            string Target = Target_saving.Text.Trim();
            string Saved = Saved_saving.Text.Trim();

            // Validate inputs
            if (string.IsNullOrWhiteSpace(date) || string.IsNullOrWhiteSpace(goal) ||
                string.IsNullOrWhiteSpace(Target) || string.IsNullOrWhiteSpace(Saved))
            {
                MessageBox.Show("Please fill in all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!decimal.TryParse(Saved, out decimal savedAmount) ||
                !decimal.TryParse(Target, out decimal targetAmount) || targetAmount <= 0)
            {
                MessageBox.Show("Please enter valid numeric values for Target and Saved.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // 1️⃣ Check if goal already exists
                    string checkQuery = "SELECT Saved, Target FROM Saving WHERE Goal = @Goal";
                    decimal existingSaved = 0;
                    decimal existingTarget = 0;
                    bool goalExists = false;

                    using (SqlCommand checkCmd = new SqlCommand(checkQuery, conn))
                    {
                        checkCmd.Parameters.AddWithValue("@Goal", goal);
                        using (SqlDataReader reader = checkCmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                existingSaved = reader["Saved"] != DBNull.Value ? Convert.ToDecimal(reader["Saved"]) : 0;
                                existingTarget = reader["Target"] != DBNull.Value ? Convert.ToDecimal(reader["Target"]) : 0;
                                goalExists = true;
                            }
                        }
                    }

                    if (goalExists)
                    {
                        // 2️⃣ Update existing goal
                        decimal newSaved = existingSaved + savedAmount;
                        decimal targetToUse = existingTarget > 0 ? existingTarget : targetAmount;
                        decimal progressPercent = (newSaved / targetToUse) * 100;
                        string progressText = progressPercent.ToString("F2") + "%"; // store as string

                        string updateQuery = @"UPDATE Saving 
                                       SET Saved = @NewSaved, Progress = @Progress
                                       WHERE Goal = @Goal";

                        using (SqlCommand updateCmd = new SqlCommand(updateQuery, conn))
                        {
                            updateCmd.Parameters.AddWithValue("@NewSaved", newSaved);
                            updateCmd.Parameters.AddWithValue("@Progress", progressText);
                            updateCmd.Parameters.AddWithValue("@Goal", goal);

                            int rows = updateCmd.ExecuteNonQuery();

                            if (rows > 0)
                            {
                                MessageBox.Show($"Updated '{goal}'. Total Saved: {newSaved}, Progress: {progressText}",
                                    "Updated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("Failed to update the existing goal.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    else
                    {
                        // 3️⃣ Insert new goal
                        decimal progressPercent = (savedAmount / targetAmount) * 100;
                        string progressText = progressPercent.ToString("F2") + "%";

                        string idQuery = "SELECT ISNULL(MAX(Id), 0) + 1 FROM Saving";
                        int nextId;
                        using (SqlCommand idCmd = new SqlCommand(idQuery, conn))
                        {
                            nextId = Convert.ToInt32(idCmd.ExecuteScalar());
                        }

                        string insertQuery = @"INSERT INTO Saving (Id, Goal, Target, Saved, Progress, Deadline)
                                       VALUES (@Id, @Goal, @Target, @Saved, @Progress, @Deadline)";

                        using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                        {
                            cmd.Parameters.AddWithValue("@Id", nextId);
                            cmd.Parameters.AddWithValue("@Goal", goal);
                            cmd.Parameters.AddWithValue("@Target", targetAmount);
                            cmd.Parameters.AddWithValue("@Saved", savedAmount);
                            cmd.Parameters.AddWithValue("@Progress", progressText);
                            cmd.Parameters.AddWithValue("@Deadline", date);

                            int rows = cmd.ExecuteNonQuery();

                            if (rows > 0)
                            {
                                MessageBox.Show($"New goal '{goal}' added successfully! Progress: {progressText}",
                                    "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("Failed to add new goal.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    LoadsavingFromDatabase(); LoadexpensessavingFromDatabase(); CreateGoalProgressChart();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}