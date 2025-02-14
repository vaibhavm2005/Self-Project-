using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Data;
using System.Data.Odbc;
using OfficeOpenXml;
using System.IO;




      

namespace demo
{


  
/// <summary>
/// Interaction logic for Window3.xaml
/// </summary>
/// 

/*public partial class Window3 : Window
{
    private DataTable dataTable = new DataTable();

    public Window3()
    {
        InitializeComponent();
        LoadExcelData();
    }
private void SaveButton_Click(object sender, RoutedEventArgs e)
{
    SaveUpdatedValueToExcel();
}

// Function to Save Updated Value to Excel
private void SaveUpdatedValueToExcel()
{
    string filePath = @"C:\Users\hp\Desktop\JAI MINESH-2024 -.xlsx"; // Correct the file path

    try
    {
        // Load the Excel file
        FileInfo file = new FileInfo(filePath);

        // Enable EPPlus to read/write to Excel
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (ExcelPackage package = new ExcelPackage(file))
        {
            // Access the first worksheet
            var worksheet = package.Workbook.Worksheets[0];

            // Update the value in F4 (row 4, column 6)
            string newValue = "45597"; // Replace this with the updated value
            worksheet.Cells[4, 6].Value = newValue;

            // Save the updated Excel file
            package.Save();
        }

        MessageBox.Show("Value updated in Excel file!");
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Error: {ex.Message}");
    }
}

    private void LoadExcelData()
    {
        string connectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=C:\\Users\\hp\\Desktop\\JAI MINESH-2024 -.xlsx;";
        using (OdbcConnection connection = new OdbcConnection(connectionString))
        {
            try
            {
                connection.Open();
                string query = "SELECT * FROM [BALANCE SHEET$]"; // Adjust sheet name
                OdbcDataAdapter adapter = new OdbcDataAdapter(query, connection);
                adapter.Fill(dataTable);
                ExcelDataGrid.ItemsSource = dataTable.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }
    }
private void UpdateMonthButton_Click(object sender, RoutedEventArgs e)
            {
                string newMonth = MonthTextBox.Text;
                if (string.IsNullOrWhiteSpace(newMonth))
                {
                    MessageBox.Show("Please enter a valid month.");
                    return;
                }

                UpdateMonthInExcel(newMonth);
            }
        private string excelFilePath = @"C:\Users\hp\Desktop\JAI MINESH-2024 -.xlsx";
            private void UpdateMonthInExcel(string newMonth)
            {
                string connectionString = $"Driver={{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}};Dbq={excelFilePath};ReadOnly=0;";
                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    try
                    {
                        connection.Open();
                        string query = $"UPDATE [BALANCE SHEET$] SET F4 = '{newMonth}' WHERE F2 = 'MONTHLY BALANCE SHEET'";
                        using (OdbcCommand command = new OdbcCommand(query, connection))
                        {
                            int rowsAffected = command.ExecuteNonQuery();
                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Month updated successfully!");
                            }
                            else
                            {
                                MessageBox.Show("No matching data found to update.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error updating Excel file: {ex.Message}");
                    }
                }
            }



private void ExcelDataGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
{

}
}
}*/





 public partial class Window3 : Window
     {
         public Window3()
         {
             InitializeComponent();
             LoadExcelData();
         }

         private void LoadExcelData()
         {
         string connectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=C:\\Users\\hp\\Desktop\\JAI MINESH-2024 -.xlsx;";
             using (OdbcConnection connection = new OdbcConnection(connectionString))
             {
                 try
                 {
                     connection.Open();
                     string query = "SELECT * FROM [BALANCE SHEET$]"; // Adjust sheet name as per your Excel file
                     OdbcDataAdapter adapter = new OdbcDataAdapter(query, connection);
                     DataTable dataTable = new DataTable();
                     adapter.Fill(dataTable);
                     ExcelDataGrid.ItemsSource = dataTable.DefaultView;
                 }
                 catch (Exception ex)
                 {
                     MessageBox.Show($"Error: {ex.Message}");
                 }
             }
         }

     private void ExcelDataGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
     {

     }
 }
 }




