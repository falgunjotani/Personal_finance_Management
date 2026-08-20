using System;
using System.Drawing;
using System.Windows.Forms;

namespace Personal_Finance_Management
{
    public partial class Form2 : Form
    {
        private Button currentButton;

        // UserControl instances for better performance
        private Dashboard dashboard;
        private Expenses expenses;
        private Income income;
        private Budgeting budgeting;
        private Investment investment;
        private Saving saving;

        public Form2()
        {
            InitializeComponent();
            InitializeUserControls();
            MakeFormResponsive();

        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            LoadControl(dashboard);  // Load Dashboard by default
            ActivateButton("Dashboard");
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Check if the user is closing the form (not programmatically)
            if (e.CloseReason == CloseReason.UserClosing)
            {
                Application.Exit(); // Exit the entire application
            }
        }

        private void MakeFormResponsive()
        {
            // Form properties
            this.Text = "Personal Finance Management";
            this.Size = new Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(900, 600);
            this.BackColor = Color.White;

            // Configure existing panels for responsiveness
            panelSidebar.Dock = DockStyle.Left;
            panelSidebar.Width = 220;
            panelSidebar.BackColor = Color.FromArgb(41, 41, 61);
            panelSidebar.Padding = new Padding(5);

            panelContent.Dock = DockStyle.Fill;
            panelContent.BackColor = Color.FromArgb(240, 240, 245);
            panelContent.Padding = new Padding(5);

            // Initialize sidebar with responsive buttons
            InitializeSidebar();

            // Handle form resize for responsiveness
            this.SizeChanged += Form2_SizeChanged;
        }

        // Dynamically add sidebar buttons
        private void InitializeSidebar()
        {
            panelSidebar.Controls.Clear();

            // Add sidebar title
            Label sidebarTitle = new Label();
            sidebarTitle.Text = "FINANCE MANAGER";
            sidebarTitle.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            sidebarTitle.ForeColor = Color.FromArgb(180, 180, 200);
            sidebarTitle.Dock = DockStyle.Top;
            sidebarTitle.Height = 50;
            sidebarTitle.TextAlign = ContentAlignment.MiddleLeft;
            sidebarTitle.Padding = new Padding(10, 0, 0, 0);
            panelSidebar.Controls.Add(sidebarTitle);

            // Add buttons in your preferred order
            AddSidebarButton("💰 Expenses", "Expenses");
            AddSidebarButton("💵 Income", "Income");
            AddSidebarButton("📊 Budgeting", "Budgeting");
            AddSidebarButton("📈 Investment", "Investment");
            AddSidebarButton("🎯 Saving", "Saving");
            AddSidebarButton("🏠 Dashboard", "Dashboard");

            // Add spacer at bottom
            Panel spacer = new Panel();
            spacer.Dock = DockStyle.Bottom;
            spacer.Height = 20;
            panelSidebar.Controls.Add(spacer);
        }

        // Helper method to add buttons
        private void AddSidebarButton(string text, string pageName)
        {
            Button btn = new Button();
            btn.Text = text;
            btn.Dock = DockStyle.Top;
            btn.Height = 50;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.BackColor = Color.FromArgb(41, 41, 61);
            btn.ForeColor = Color.Gainsboro;
            btn.Font = new Font("Segoe UI", 10, FontStyle.Regular);
            btn.TextAlign = ContentAlignment.MiddleLeft;
            btn.Padding = new Padding(15, 0, 0, 0);
            btn.Margin = new Padding(0, 1, 0, 1);
            btn.Cursor = Cursors.Hand;
            btn.Tag = pageName;

            // Click event
            btn.Click += (s, e) =>
            {
                LoadPage(pageName);
                ActivateButton(pageName);
            };

            // Hover effects
            btn.MouseEnter += (s, e) =>
            {
                if (btn != currentButton)
                {
                    btn.BackColor = Color.FromArgb(60, 60, 85);
                    btn.ForeColor = Color.White;
                }
            };

            btn.MouseLeave += (s, e) =>
            {
                if (btn != currentButton)
                {
                    btn.BackColor = Color.FromArgb(41, 41, 61);
                    btn.ForeColor = Color.Gainsboro;
                }
            };

            panelSidebar.Controls.Add(btn);
        }

        private void InitializeUserControls()
        {
            // Initialize all user controls once
            dashboard = new Dashboard();
            expenses = new Expenses();
            income = new Income();
            budgeting = new Budgeting();
            investment = new Investment();
            saving = new Saving();
        }

        private void LoadPage(string pageName)
        {
            UserControl control;

            // Traditional switch statement for C# 7.3 compatibility
            switch (pageName)
            {
                case "Dashboard":
                    // Create new instance to reload
                    dashboard = new Dashboard();
                    control = dashboard;
                    break;
                case "Expenses":
                    // Create new instance to reload
                    expenses = new Expenses();
                    control = expenses;
                    break;
                case "Income":
                    // Create new instance to reload
                    income = new Income();
                    control = income;
                    break;
                case "Budgeting":
                    // Create new instance to reload
                    budgeting = new Budgeting();
                    control = budgeting;
                    break;
                case "Investment":
                    // Create new instance to reload
                    investment = new Investment();
                    control = investment;
                    break;
                case "Saving":
                    // Create new instance to reload
                    saving = new Saving();
                    control = saving;
                    break;
                default:
                    dashboard = new Dashboard();
                    control = dashboard;
                    break;
            }

            LoadControl(control);
        }

        // Alternative approach: Use Refresh methods in each UserControl
        private void LoadPageWithRefresh(string pageName)
        {
            UserControl control = null;

            switch (pageName)
            {
                case "Dashboard":
                    control = dashboard;
                    if (control is Dashboard dashboardControl)
                    {
                        // Call refresh method if it exists
                        dashboardControl.RefreshData();
                    }
                    break;
                case "Expenses":
                    control = expenses;
                    if (control is Expenses expensesControl)
                    {
                        expensesControl.RefreshData();
                    }
                    break;
                case "Income":
                    control = income;
                    if (control is Income incomeControl)
                    {
                        incomeControl.RefreshData();
                    }
                    break;
                case "Budgeting":
                    control = budgeting;
                    if (control is Budgeting budgetingControl)
                    {
                        budgetingControl.RefreshData();
                    }
                    break;
                case "Investment":
                    control = investment;
                    if (control is Investment investmentControl)
                    {
                        investmentControl.RefreshData();
                    }
                    break;
                case "Saving":
                    control = saving;
                    if (control is Saving savingControl)
                    {
                        savingControl.RefreshData();
                    }
                    break;
                default:
                    control = dashboard;
                    break;
            }

            if (control != null)
            {
                LoadControl(control);
            }
        }

        // Loads the given UserControl into panelContent
        private void LoadControl(UserControl control)
        {
            panelContent.SuspendLayout();
            panelContent.Controls.Clear();

            if (control != null)
            {
                control.Dock = DockStyle.Fill;
                panelContent.Controls.Add(control);
            }

            panelContent.ResumeLayout();
        }

        private void ActivateButton(string pageName)
        {
            // Deactivate current button
            if (currentButton != null)
            {
                currentButton.BackColor = Color.FromArgb(41, 41, 61);
                currentButton.ForeColor = Color.Gainsboro;
                currentButton.Font = new Font(currentButton.Font, FontStyle.Regular);
            }

            // Find and activate new button
            foreach (Control control in panelSidebar.Controls)
            {
                if (control is Button btn && btn.Tag?.ToString() == pageName)
                {
                    currentButton = btn;
                    btn.BackColor = Color.FromArgb(0, 150, 136);
                    btn.ForeColor = Color.White;
                    btn.Font = new Font(btn.Font, FontStyle.Bold);
                    break;
                }
            }
        }

        private void Form2_SizeChanged(object sender, EventArgs e)
        {
            UpdateResponsiveLayout();
        }

        private void UpdateResponsiveLayout()
        {
            try
            {
                // Adjust layout based on form width
                if (this.Width < 1000)
                {
                    // Small screen layout
                    ApplySmallScreenLayout();
                }
                else if (this.Width < 1300)
                {
                    // Medium screen layout
                    ApplyMediumScreenLayout();
                }
                else
                {
                    // Large screen layout
                    ApplyLargeScreenLayout();
                }
            }
            catch (Exception ex)
            {
                // Safe fallback for layout errors
                Console.WriteLine($"Layout update error: {ex.Message}");
            }
        }

        private void ApplySmallScreenLayout()
        {
            // Compact sidebar for small screens
            panelSidebar.Width = 180;

            // Smaller fonts
            UpdateSidebarFonts(9);
        }

        private void ApplyMediumScreenLayout()
        {
            // Medium sidebar width
            panelSidebar.Width = 220;

            // Medium fonts
            UpdateSidebarFonts(10);
        }

        private void ApplyLargeScreenLayout()
        {
            // Larger sidebar width
            panelSidebar.Width = 240;

            // Normal fonts
            UpdateSidebarFonts(10);
        }

        private void UpdateSidebarFonts(int fontSize)
        {
            foreach (Control control in panelSidebar.Controls)
            {
                if (control is Button btn)
                {
                    btn.Font = new Font("Segoe UI", fontSize, btn.Font.Style);
                }
                else if (control is Label label)
                {
                    label.Font = new Font("Segoe UI", fontSize, FontStyle.Bold);
                }
            }
        }
    }
}