using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
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
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public static List<MyTable> compList = new List<MyTable>();
        public int mult = 0;
        public string fPath = Environment.CurrentDirectory + "/" + "thrlist.xlsx";
        
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            CheckTheList();
            
        }

        private void CheckTheList()
        {
            if (File.Exists(fPath))
            {
                MessageBox.Show("База содержится на компьютере");
            }
            else
            {
                AskWindow ask = new AskWindow();
                if (ask.ShowDialog() == true)
                {
                    DownloadFile();
                }
                else
                {
                    MessageBox.Show("До встречи!");
                    System.Windows.Application.Current.Shutdown();
                }
            }
        }
        public void CreateList(int mult)
        {
            Excel excel = new Excel(Environment.CurrentDirectory + "/" + "thrlist.xlsx", 1);
            for (int i = 1; i < 16; i++)
            {
                MyTable shortTable = new MyTable();
                {
                    shortTable.Ubi = "УБИ." + excel.ReadCell(i + (15 * mult) + 1, 0).ToString();
                    shortTable.Name = excel.ReadCell(i + (15  * mult) + 1, 1).ToString();
                    shortTable.Id = excel.ReadCell(i + (15  * mult) + 1, 0).ToString();
                    shortTable.Description = excel.ReadCell(i + (15  * mult) + 1, 2).ToString();
                    shortTable.Source = excel.ReadCell(i + (15  * mult) + 1, 3).ToString();
                    shortTable.ObjectVozd = excel.ReadCell(i + (15  * mult) + 1, 4).ToString();
                    shortTable.Conf = excel.ReadCell(i + (15  * mult) + 1, 5).ToString();
                    shortTable.Celostn = excel.ReadCell(i + (15  * mult) + 1, 6).ToString();
                    shortTable.Access = excel.ReadCell(i + (15  * mult) + 1, 7).ToString();
                    if (excel.ReadCellDate(i + (15 * mult) + 1, 8) != "blank")
                    {
                        shortTable.Date = new DateTime(1899, 12, 30).AddDays(Convert.ToDouble(excel.ReadCellDate(i + (15 * mult) + 1, 8))).ToString();
                    }
                    else shortTable.Date = "blank";
                    if (excel.ReadCellDate(i + (15 * mult) + 1, 8) != "blank")
                    {
                        shortTable.LastChange = new DateTime(1899, 12, 30).AddDays(Convert.ToDouble(excel.ReadCellDate(i + (15 * mult) + 1, 9))).ToString();
                    }
                    else shortTable.Date = "blank";
                    
                }
                listName.Items.Add(shortTable);
            }
            excel.Close();
        }

        public void CopyFile(string name)
        {
            _Excel.Application xlAp = new _Excel.Application();
            _Excel.Workbook wbAp;
            _Excel.Worksheet wsAp;
            object MisValue = System.Reflection.Missing.Value;
            wbAp = xlAp.Workbooks.Add(MisValue);
            wsAp = (_Excel.Worksheet)wbAp.Worksheets.get_Item(1);
            try
            {
                wbAp.SaveAs(Environment.CurrentDirectory + "/" + name);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Тут должно было быть исключение");
            }
            wbAp.Close(true, MisValue, MisValue);
            xlAp.Quit();
            Marshal.ReleaseComObject(wsAp);
            Marshal.ReleaseComObject(wbAp);
            Marshal.ReleaseComObject(xlAp);

            try
            {
                File.Copy(Environment.CurrentDirectory + "/" + "thrlist.xlsx", Environment.CurrentDirectory + "/" + name, true);
                MessageBox.Show("Файл сохранен в директории программы под именем " + name);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        public void DownloadFile()
        {
            WebClient wc = new WebClient();
            string url = "https://bdu.fstec.ru/files/documents/thrlist.xlsx";
            string save_path = Environment.CurrentDirectory + "/" + "thrlist.xlsx";
            string name = "thrlist.xlsx";
            try
            {
                wc.DownloadFile(url, save_path);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка загрузки");
            }
            if (File.Exists(fPath))
            {
                MessageBox.Show("Загрузка прошла успешно");
            }
            else
            {
                MessageBox.Show("Ошибка загрузки!");
            }
        }

        public void CompareLists()
        {
            compList.Clear();
            string ex1 = Environment.CurrentDirectory + "/" + "thrlist.xlsx";
            string ex2 = Environment.CurrentDirectory + "/" + "prevExcel.xlsx";

            _Excel.Application app = new _Excel.Application();

            try
            {
                Workbook wb1 = app.Workbooks.Open(ex1);
                Workbook wb2 = app.Workbooks.Open(ex2);

                Worksheet sh1 = wb1.Sheets[1];
                Worksheet sh2 = wb2.Sheets[1];
                Range rng1 = sh1.Range["A1", "J240"];
                Range rng2 = sh2.Range["A1", "J240"];

                for (int i = 1; i <= rng1.Rows.Count; i++)
                {
                    for (int j = 1; j <= rng1.Columns.Count; j++)
                    {
                        var val1 = rng1.Cells[i, j].Value;
                        var val2 = rng2.Cells[i, j].Value;

                        if (val1 != val2)
                        {
                            Excel excel = new Excel(Environment.CurrentDirectory + "/" + "thrlist.xlsx", 1);
                            Excel excel1 = new Excel(Environment.CurrentDirectory + "/" + "prevExcel.xlsx", 1);
                            MyTable compTable = new MyTable();
                            {
                                compTable.Id = excel.ReadCell(i + 1, 0).ToString();
                                compTable.InPast = excel.ReadCell(i + 1, j).ToString();
                                compTable.Today = excel1.ReadCell(i + 1, j).ToString();
                            }
                            compList.Add(compTable);
                            rng2.Cells[i, j].Interior.Color = 255;
                        }
                    }
                }

                app.DisplayAlerts = false;

                wb1.Close(true);
                wb2.Close(true);

                Window1 window1 = new Window1();
                window1.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        
        private void listName_Loaded(object sender, RoutedEventArgs e)
        {
            CreateList(0);
        }

        private void buttonNext_Click(object sender, RoutedEventArgs e)
        {
            listName.Items.Clear();
            mult++;
            CreateList(mult);
        }

        private void buttonPrev_Click(object sender, RoutedEventArgs e)
        {
            if(mult > 0)
            {
                listName.Items.Clear();
                mult--;
                CreateList(mult);
            }        
        }

        private void listName_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
           //MessageBox.Show("asd");
        }

        private void buttonCreate_Click(object sender, RoutedEventArgs e)
        {
            CopyFile("newExcel.xlsx");

        }

        private void buttonUpload_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("В файле PrevExcel для удобства будут помечены красным изменненные поля");
            CopyFile("prevExcel.xlsx");
            DownloadFile();
            listName.Items.Clear();
            CreateList(0);
            CompareLists();
        }



    }
}
