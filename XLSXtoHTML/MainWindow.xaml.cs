using System;
using System.Collections.Generic;
using System.ComponentModel;
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

namespace XLSXtoHTML
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string Load(string filePath)
        {
            string result = "";

            Excel.Application app = new();
            Excel.Workbook workbook = app.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = (Excel.Worksheet) workbook.ActiveSheet;

            Excel.Range usedRange = worksheet.UsedRange;

            for (int i = 1; i <= usedRange.Columns.Count; i++)
            {
                for (int j = 1; j <= usedRange.Rows.Count; j++)
                {
                    Excel.Range CellRange = usedRange.Cells[i, j] as Excel.Range;
                    // Получение текста ячейки
                    string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                        (CellRange as Excel.Range).Value2.ToString();

                    if (CellText != null)
                    {
                        result += CellText;
                    }

                    if (j != usedRange.Columns.Count)
                    {
                        result += "\t";
                    }
                    else if (i != usedRange.Rows.Count)
                    {
                        result += "\n";
                    }
                }
            }

            workbook.Close(false);
            app.Quit();

            return result;
        }

        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            ResultWindow resultWindow = new();
            resultWindow.LoadResult(Load("C:\\Users\\dmitr\\Documents\\lab\\XLSXtoDOC\\excel.xlsx"));
            resultWindow.Show();
        }
    }
}
