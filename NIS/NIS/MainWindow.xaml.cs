using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
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

namespace NIS
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        string fileName;
        string[] file;
        List<Row> rows = new List<Row>();
        double allPoints;

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MyWindow_Loaded;
        }
        private void MyWindow_Loaded(object sender, RoutedEventArgs e)
        {

        }
        private List<string> createCriteria(string[] line)
        {
            List<string> listToResult = new List<string>();
            for (int i = 3; i < line.Length; i++)
            {
                listToResult.Add(line[i]);
            }
            return listToResult;
        }
        void ConvertExcelToCsv(string excelFilePath, string csvOutputFile, int worksheetNumber = 1)
        {
            if (!File.Exists(excelFilePath)) MessageBox.Show("Выберите файл в следующий раз");
            if (File.Exists(csvOutputFile)) File.Delete(csvOutputFile);

            // connection string
            var cnnStr = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO\"", excelFilePath);
            var cnn = new OleDbConnection(cnnStr);

            // get schema, then data
            var dt = new System.Data.DataTable();
            try
            {
                cnn.Open();
                var schemaTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schemaTable.Rows.Count < worksheetNumber) throw new ArgumentException("The worksheet number provided cannot be found in the spreadsheet");
                string worksheet = schemaTable.Rows[worksheetNumber - 1]["table_name"].ToString().Replace("'", "");
                string sql = String.Format("select * from [{0}]", worksheet);
                var da = new OleDbDataAdapter(sql, cnn);
                da.Fill(dt);
            }
            catch (Exception e)
            {
                // ???
                throw e;
            }
            finally
            {
                // free resources
                cnn.Close();
            }

            // write out CSV data
            using (var wtr = new StreamWriter(csvOutputFile, false, Encoding.UTF8))
            {
                foreach (DataRow row in dt.Rows)
                {
                    bool firstLine = true;
                    foreach (DataColumn col in dt.Columns)
                    {
                        if (!firstLine) { wtr.Write(","); } else { firstLine = false; }
                        var data = row[col.ColumnName].ToString().Replace("\"", "\"\"");
                        wtr.Write(String.Format("\"{0}\"", data));
                    }
                    wtr.WriteLine();
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Файлы Excel (*.xls; *.xlsx) | *.xls; *.xlsx";
            openFileDialog.ShowDialog();
            fileName = openFileDialog.FileName;

            ConvertExcelToCsv(fileName, @"test.csv");
            file = File.ReadAllLines(@"test.csv", Encoding.UTF8);

            foreach (string row in file)
            {
                string rowInside = row.Replace("\"", "");
                string[] line = rowInside.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                int id = 0;
                if (line.Length < 2 ||
                    !Int32.TryParse(line[0], out id))
                {
                    continue;
                }
                rows.Add(new Row
                {
                    Id = Int32.Parse(line[0]),
                    Manager = line[1],
                    Name = line[2],
                    Criteria = createCriteria(line)
                });
            }

            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new
                Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }
                object misValue = System.Reflection.Missing.Value;
                Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);

                Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "Имя";
                xlWorkSheet.Cells[1, 2] = "Проект";
                xlWorkSheet.Cells[1, 3] = "Абсолютный балл";
                xlWorkSheet.Cells[1, 4] = "Относительный балл";
                for (int i = 0; i < rows.Count; i++)
                {
                    xlWorkSheet.Cells[i + 2, 1] = rows[i].Manager;
                    xlWorkSheet.Cells[i + 2, 2] = rows[i].Name;
                    xlWorkSheet.Cells[i + 2, 3] = returnPoint(i);
                }
                for (int i = 0; i < rows.Count; i++)
                {
                    xlWorkSheet.Cells[i + 2, 4] = returnPercents(i);
                }

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
                xlWorkBook.SaveAs(@"../../wwdwd.xls", XlFileFormat.xlExcel9795, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlShared, misValue, misValue, misValue, misValue, misValue);
            }
            catch (Exception) { }
            File.Delete(@"test.csv");
        }
        private string returnPercents(int i)
        {
            return String.Format("{0}%", Math.Round((rows[i].Sum / allPoints) * 100, 2));
        }
        private double returnPoint(int i)
        {
            var list = rows[i].Criteria;
            double result = 0;
            foreach (var item in list)
            {

                switch (item)
                {
                    case "НЕТ":
                        result += 0;
                        break;
                    case "СЛАБО":
                        result += 1;
                        break;
                    case "СРЕДНЕ":
                        result += 2;
                        break;
                    case "СИЛЬНО":
                        result += 3;
                        break;
                    case "ПОЛНОСТЬЮ":
                        result += 4;
                        break;
                    default:
                        break;
                }
            }
            allPoints += Math.Round(result / list.Count, 2);
            rows[i].Sum = Math.Round(result / list.Count, 2);
            return Math.Round(result / list.Count, 2);
        }
    }
}
