using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Personal_Finance_Management
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.MaximizeBox = false;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Check if the user is closing the form (not programmatically)
            if (e.CloseReason == CloseReason.UserClosing)
            {
                Application.Exit(); // Exit the entire application
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string username = Username.Text.Trim();
            string password = Password.Text.Trim();

            if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Please enter both username and password.", "Validation Error",
                               MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string connectionString = AppData.ConnectionString;
            // Try different connection string formats
           

            // Correct query - Table: logintable
            string query = "SELECT COUNT(*) FROM logintable WHERE username = @Username AND password = @Password";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open(); // Try opening connection first to test connectivity
                   

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Username", username);
                        command.Parameters.AddWithValue("@Password", password);

                        int result = (int)command.ExecuteScalar();

                        if (result > 0)
                        {
                            Form2 form2 = new Form2();
                            form2.Show();
                            this.Hide();
                        }
                        else
                        {
                            MessageBox.Show("Username and password don't match.", "Login Error",
                                           MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Password.Clear();
                            Username.Focus();
                        }
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                MessageBox.Show($"SQL Error: {sqlEx.Message}\nError Number: {sqlEx.Number}", "Database Error",
                               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error",
                               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

       
    }
}