# Personal Finance Management App (Desktop)

Personal Finance Manager built with C# WinForms - Track income, expenses, investments, and savings. Includes responsive dashboard, color-coded transactions, and SQL Server backend for robust data management.

## Quick Start

1. **Restore Database**: 
   - Open SQL Server Management Studio
   - Restore `Database/personal_finance.bak` file

2. **Open in Visual Studio**:
   - Open the project solution file
   - Build and run the application
  
     
3.  **change the connectionstring**:
     -  string connectionString = @"Data Source=Server_name\SQLEXPRESS;Initial Catalog=personal_finance;Integrated Security=True;";
In "Server_name\SQLEXPRESS" use your Server name
  
4. **Login**:
   - Username: `admin`
   - Password: `admin`
   - (Credentials are pre-filled in the login form)

After login, the dashboard will appear with your financial overview.
