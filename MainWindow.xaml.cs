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
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
namespace Scouting
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            getExcelFile();
            //dataProccessing();
        }

        public static void dataProccessing()
        {

        }

        public static void getExcelFile()
        {
            StreamWriter sw = new StreamWriter(/*file path for output txt file*/);
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(/*file path for excel file: add excel file as reference to project*/);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for(int i = 1; i <= rowCount; i++)
            {
                for(int j = 1; j <= colCount; j++)
                {
                    if( j == 1)
                    {
                        sw.Write("\r\n");
                    }
                    if(xlRange.Cells[i,j] != null && xlRange.Cells[i,j].Value2 != null)
                    {
                        sw.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                    }
                }
            }
            sw.Flush();
            sw.Close();
            MessageBox.Show("Excel copying complete");
        }
    }
}
