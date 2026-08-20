using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace Personal_Finance_Management
{
    public partial class Dashboard : UserControl
    {
        private TableLayoutPanel mainLayout;
        private FlowLayoutPanel cardsFlowPanel;
        private DataGridView transactionsGrid;
        private Panel contentPanel;
        private Label transactionsHeader;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;

        string connectionString = AppData.ConnectionString;
        public Dashboard()
        {
            InitializeComponent();
            CreateResponsiveDashboard();
            LoadTransactionData();
        }

        // ---------------------- UI CREATION ----------------------
        private void CreateResponsiveDashboard()
        {
            this.BackColor = Color.FromArgb(240, 240, 245);
            this.Padding = new Padding(20);
            this.Dock = DockStyle.Fill;

            mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10)
            };

            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 80));   // Header
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 150));  // Summary Cards
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));   // Content Area

            CreateHeader(mainLayout);
            CreateSummaryCards(mainLayout);
            CreateContentArea(mainLayout);

            this.Controls.Add(mainLayout);
            this.SizeChanged += Dashboard_SizeChanged;
        }

        private void CreateHeader(TableLayoutPanel parent)
        {
            Label title = new Label
            {
                Text = "Dashboard Overview",
                Font = new Font("Segoe UI", 24, FontStyle.Bold),
                ForeColor = Color.FromArgb(51, 51, 76),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            };

            Panel headerPanel = new Panel { Dock = DockStyle.Fill };
            headerPanel.Controls.Add(title);
            parent.Controls.Add(headerPanel, 0, 0);
        }

        private void CreateSummaryCards(TableLayoutPanel parent)
        {
            cardsFlowPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                WrapContents = true,
                Padding = new Padding(10)
            };

            string[] titles = { "Total Investment", "Monthly Income", "Monthly Expenses", "Savings Rate" };

            string savingsRate = GetAverageProgress();
            decimal totalinvestment = GetTotalInvestment();
            decimal monthlyIncome = GetCurrentMonthIncome();
            decimal monthlyExpense = GetCurrentMonthExpense();

            string[] values = {
                "¥" + totalinvestment.ToString("N2"),
                "¥" + monthlyIncome.ToString("N2"),
                "¥" + monthlyExpense.ToString("N2"),
                savingsRate
            };

            Color[] colors = {
                Color.FromArgb(0, 150, 136),
                Color.FromArgb(76, 175, 80),
                Color.FromArgb(244, 67, 54),
                Color.FromArgb(156, 39, 176)
            };

            for (int i = 0; i < titles.Length; i++)
            {
                Panel card = CreateSummaryCard(titles[i], values[i], colors[i]);
                cardsFlowPanel.Controls.Add(card);
            }

            parent.Controls.Add(cardsFlowPanel, 0, 1);
        }

        private Panel CreateSummaryCard(string title, string value, Color color)
        {
            Panel card = new Panel
            {
                Size = new Size(220, 120),
                Margin = new Padding(5),
                BackColor = color,
                Padding = new Padding(15),
                Cursor = Cursors.Hand
            };

            card.MouseEnter += (s, e) => { card.BackColor = ControlPaint.Dark(color, 0.1f); };
            card.MouseLeave += (s, e) => { card.BackColor = color; };

            Label titleLabel = new Label
            {
                Text = title,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10),
                AutoSize = true,
                Location = new Point(15, 15)
            };

            Label valueLabel = new Label
            {
                Text = value,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 18, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(15, 45)
            };

            card.Controls.Add(titleLabel);
            card.Controls.Add(valueLabel);
            return card;
        }

        private void CreateContentArea(TableLayoutPanel parent)
        {
            contentPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };

            transactionsHeader = new Label
            {
                Text = "Recent Transactions",
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = Color.FromArgb(51, 51, 76),
                Location = new Point(20, 20),
                Size = new Size(300, 30),
                AutoSize = true
            };

            CreateTransactionsGrid();

            contentPanel.Controls.Add(transactionsHeader);
            contentPanel.Controls.Add(transactionsGrid);

            parent.Controls.Add(contentPanel, 0, 2);
        }

        private void CreateTransactionsGrid()
        {
            transactionsGrid = new DataGridView
            {
                Location = new Point(20, 60), // Moved up since no buttons
                Size = new Size(800, 400),
                BorderStyle = BorderStyle.None,
                BackgroundColor = Color.White,
                GridColor = Color.FromArgb(224, 224, 224),
                Font = new Font("Segoe UI", 10),
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            transactionsGrid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(51, 51, 76);
            transactionsGrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            transactionsGrid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            transactionsGrid.ColumnHeadersHeight = 40;
            transactionsGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 248, 248);
            transactionsGrid.RowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(33, 150, 243);
            transactionsGrid.RowsDefaultCellStyle.SelectionForeColor = Color.White;
        }

        // ---------------------- DATA LOADING ----------------------
        private void LoadTransactionData()
        {
            try
            {
                DataTable data = GetUnifiedTransactionData();

                if (data.Rows.Count > 0)
                {
                    transactionsGrid.DataSource = data;

                    if (transactionsGrid.Columns.Contains("Amount"))
                    {
                        transactionsGrid.Columns["Amount"].DefaultCellStyle.Format = "C2";
                        transactionsGrid.Columns["Amount"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }

                    // Color code rows
                    transactionsGrid.RowPrePaint += (s, e) =>
                    {
                        if (e.RowIndex >= 0 && e.RowIndex < transactionsGrid.Rows.Count)
                        {
                            var row = transactionsGrid.Rows[e.RowIndex];
                            if (row.Cells["Type"]?.Value != null)
                            {
                                string type = row.Cells["Type"].Value.ToString();
                                if (type == "Income")
                                {
                                    row.DefaultCellStyle.ForeColor = Color.FromArgb(76, 175, 80);
                                    row.DefaultCellStyle.SelectionForeColor = Color.White;
                                }
                                else if (type == "Expense")
                                {
                                    row.DefaultCellStyle.ForeColor = Color.FromArgb(244, 67, 54);
                                    row.DefaultCellStyle.SelectionForeColor = Color.White;
                                }
                            }
                        }
                    };
                }
                else
                {
                    ShowNoDataMessage();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading transaction data: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowErrorMessage(ex.Message);
            }
        }

        private void ShowNoDataMessage()
        {
            transactionsGrid.DataSource = null;
            transactionsGrid.Rows.Clear();
            transactionsGrid.Columns.Clear();

            transactionsGrid.Columns.Add("Message", "Information");
            transactionsGrid.Columns["Message"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            transactionsGrid.Rows.Add("No transactions found for the current month.");
            transactionsGrid.Rows[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            transactionsGrid.Rows[0].DefaultCellStyle.ForeColor = Color.Gray;
            transactionsGrid.Rows[0].DefaultCellStyle.Font = new Font("Segoe UI", 12, FontStyle.Italic);
        }

        private void ShowErrorMessage(string message)
        {
            transactionsGrid.DataSource = null;
            transactionsGrid.Rows.Clear();
            transactionsGrid.Columns.Clear();
            transactionsGrid.Columns.Add("Error", "Error");
            transactionsGrid.Columns["Error"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            transactionsGrid.Rows.Add($"Error loading data: {message}");
            transactionsGrid.Rows[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            transactionsGrid.Rows[0].DefaultCellStyle.ForeColor = Color.Red;
        }

        private DataTable GetUnifiedTransactionData()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Type", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));
            table.Columns.Add("Description", typeof(string));
            table.Columns.Add("Category", typeof(string));
            table.Columns.Add("Amount", typeof(decimal));
            table.Columns.Add("PaymentMethod", typeof(string));

            string incomeQuery = @"
                SELECT 
                    'Income' AS Type,
                    Next_payment AS Date,
                    Source AS Description,
                    Type AS Category,
                    Amount,
                    'N/A' AS PaymentMethod
                FROM Income_sources
                WHERE MONTH(Next_payment) = MONTH(GETDATE())
                  AND YEAR(Next_payment) = YEAR(GETDATE())";

            string expenseQuery = @"
                SELECT 
                    'Expense' AS Type,
                    Date,
                    Description,
                    Category,
                    Amount,
                    Payment_method AS PaymentMethod
                FROM Recent_expenses
                WHERE MONTH(Date) = MONTH(GETDATE())
                  AND YEAR(Date) = YEAR(GETDATE())";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                // Income
                using (SqlCommand cmd = new SqlCommand(incomeQuery, conn))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        table.Rows.Add(
                            "Income",
                            reader["Date"],
                            reader["Description"],
                            reader["Category"],
                            reader["Amount"],
                            reader["PaymentMethod"]
                        );
                    }
                }

                // Expense
                using (SqlCommand cmd = new SqlCommand(expenseQuery, conn))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        table.Rows.Add(
                            "Expense",
                            reader["Date"],
                            reader["Description"],
                            reader["Category"],
                            reader["Amount"],
                            reader["PaymentMethod"]
                        );
                    }
                }
            }

            table.DefaultView.Sort = "Date DESC";
            return table.DefaultView.ToTable();
        }

        // ---------------------- RESPONSIVE LAYOUT ----------------------
        private void Dashboard_SizeChanged(object sender, EventArgs e)
        {
            UpdateResponsiveLayout();
        }

        private void UpdateResponsiveLayout()
        {
            try
            {
                // Update card sizes
                if (cardsFlowPanel != null && cardsFlowPanel.Width > 0)
                {
                    int availableWidth = cardsFlowPanel.Width - 40;
                    int cardWidth = Math.Max(200, (availableWidth - 60) / 4);

                    foreach (Control control in cardsFlowPanel.Controls)
                    {
                        if (control is Panel card)
                        {
                            card.Size = new Size(cardWidth, 120);
                        }
                    }
                }

                // Update DataGridView size and position
                if (contentPanel != null && transactionsGrid != null)
                {
                    int contentWidth = contentPanel.Width;
                    int contentHeight = contentPanel.Height;

                    // Responsive margins (5% of width or minimum 20px)
                    int horizontalMargin = Math.Max(20, (int)(contentWidth * 0.05));
                    int verticalMargin = Math.Max(20, (int)(contentHeight * 0.05));

                    // Calculate available space for DataGridView
                    int gridWidth = contentWidth - (2 * horizontalMargin);
                    int gridHeight = contentHeight - (transactionsHeader.Bottom + verticalMargin);

                    // Ensure minimum sizes
                    gridWidth = Math.Max(400, gridWidth);
                    gridHeight = Math.Max(200, gridHeight);

                    // Update DataGridView position and size
                    transactionsGrid.Location = new Point(horizontalMargin, transactionsHeader.Bottom + 10);
                    transactionsGrid.Size = new Size(gridWidth, gridHeight);

                    // Adjust header position
                    transactionsHeader.Location = new Point(horizontalMargin, verticalMargin);
                }
            }
            catch (Exception ex)
            {
                // Silent fail for layout issues
                System.Diagnostics.Debug.WriteLine($"Layout update error: {ex.Message}");
            }
        }

        // ---------------------- DATABASE HELPERS ----------------------
        private decimal GetTotalInvestment()
        {
            string query = "SELECT SUM(CAST(Amount_invested AS DECIMAL(18,2))) FROM Portfolio";
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                conn.Open();
                object result = cmd.ExecuteScalar();
                return (result == DBNull.Value) ? 0 : Convert.ToDecimal(result);
            }
        }

        private string GetAverageProgress()
        {
            string query = @"
                SELECT AVG(CAST(REPLACE(Progress, '%', '') AS DECIMAL(18,2))) 
                FROM Saving WHERE Progress IS NOT NULL AND Progress != ''";

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                conn.Open();
                object result = cmd.ExecuteScalar();
                if (result == DBNull.Value) return "0%";
                return Math.Round(Convert.ToDecimal(result), 2) + "%";
            }
        }

        private decimal GetCurrentMonthIncome()
        {
            string query = @"
                SELECT SUM(CAST(Amount AS DECIMAL(18,2))) 
                FROM Income_sources
                WHERE MONTH(Next_payment) = MONTH(GETDATE())
                  AND YEAR(Next_payment) = YEAR(GETDATE())";

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                conn.Open();
                object result = cmd.ExecuteScalar();
                return (result == DBNull.Value) ? 0 : Convert.ToDecimal(result);
            }
        }

        private decimal GetCurrentMonthExpense()
        {
            string query = @"
                SELECT SUM(CAST(Amount AS DECIMAL(18,2))) 
                FROM Recent_expenses
                WHERE MONTH(Date) = MONTH(GETDATE())
                  AND YEAR(Date) = YEAR(GETDATE())";

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                conn.Open();
                object result = cmd.ExecuteScalar();
                return (result == DBNull.Value) ? 0 : Convert.ToDecimal(result);
            }
        }

        private void InitializeComponent()
        {
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // Dashboard
            // 
            this.Name = "Dashboard";
            this.Size = new System.Drawing.Size(875, 788);
           // this.Load += new System.EventHandler(this.Dashboard_Load);
            this.ResumeLayout(false);

        }

        public void RefreshData()
        {
            GetCurrentMonthExpense(); GetCurrentMonthIncome(); GetAverageProgress(); GetTotalInvestment();
        }

    }
}