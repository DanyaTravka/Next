using PaymentExampleApp.Model;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel =Microsoft.Office.Interop.Excel;

namespace PaymentExampleApp.View.Pages
{
    /// <summary>
    /// Логика взаимодействия для Page1.xaml
    /// </summary>
    public partial class MainPage : Page
    {
        Core db = new Core();
        Excel.Application application;
        public MainPage()
        {
            InitializeComponent();
        }

        private void ReportButtonClick(object sender, RoutedEventArgs e)
        {
           
            //Запускаем приложение
             application = new Excel.Application();
            application.Visible = true;
            //Создание файла
            Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);
            //Формируем массив
            var allUsers = db.context.Users.OrderBy(p => p.LastName).ToList();
            //Количество листов в книге
            application.SheetsInNewWorkbook = allUsers.Count();
            
            for (int i = 0; i < allUsers.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = application.Worksheets.Item[i + 1];
                worksheet.Name = allUsers[i].LastName;
                startRowIndex++;
                //вывод заголовков
                worksheet.Cells[1][startRowIndex] = "Дата платежа";
                worksheet.Cells[2][startRowIndex] = "Название";
                worksheet.Cells[3][startRowIndex] = "Стоимость";
                worksheet.Cells[4][startRowIndex] = "Количество";
                worksheet.Cells[5][startRowIndex] = "Сумма";

                var usersCategories = allUsers[i].Pay.OrderBy(p => p.Date_payment).GroupBy(p => p.category).OrderBy(p => p.Key.Name_category);

            foreach (var groupCategory in usersCategories)
            {
                Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[5][startRowIndex]];
                headerRange.Merge();
                headerRange.Value = groupCategory.Key.Name_category;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Italic = true;

                    startRowIndex++;

                foreach (var payments in groupCategory)
                {
                        worksheet.Cells[1][startRowIndex] = "Дата платежа";
                        worksheet.Cells[2][startRowIndex] = "Название";
                        worksheet.Cells[3][startRowIndex] = "Стоимость";
                        worksheet.Cells[4][startRowIndex] = "Количество";
                        worksheet.Cells[5][startRowIndex] = "Сумма";
                }

            }
            
            }

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
           
        }
    }
}
