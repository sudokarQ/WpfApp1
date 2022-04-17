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
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public int mult = 0;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
           
        }
        public void CreateList(int mult)
        {
            Excel excel = new Excel(@"C:\C++ study\WpfApp1\WpfApp1\sheet\thrlist.xlsx", 1);
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
                    shortTable.Date = excel.ReadCell(i + (15  * mult) + 1, 8);
                    shortTable.LastChange = excel.ReadCell(i + (15 * mult) + 1, 9);
                }
                listName.Items.Add(shortTable);
            }
            excel.Close();
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
            
            
        }

        //private void listName_ItemActivate(Object sender, EventArgs e)
        //{

        //    MessageBox.Show("You are in the ListView.ItemActivate event.");
        //}

        //private void grid_Loaded(object sender, RoutedEventArgs e)
        //{
        //    Excel excel = new Excel(@"C:\C++ study\WpfApp1\WpfApp1\sheet\thrlist.xlsx", 1);
        //    List<MyTable> result = new List<MyTable>();
        //    result.Add(excel.ReadRow(3, 1, 10));
        //    result.Add(excel.ReadRow(4, 1, 10));
        //    result.Add(excel.ReadRow(5, 1, 10));
        //    result.Add(excel.ReadRow(6, 1, 10));
        //    result.Add(excel.ReadRow(7, 1, 10));
        //    //string[,] result = excel.ReadRange(2, 1, 8, 10);
        //    //result.Add(new MyTable(1, "Майкл Джексон", "Thriller", 1982));
        //    //result.Add(new MyTable(2, "AC/DC", "Back in Black", 1980));
        //    //result.Add(new MyTable(3, "Bee Gees", "Saturday Night Fever", 1977));
        //    //result.Add(new MyTable(4, "Pink Floyd", "The Dark Side of the Moon", 1973));
        //    grid.ItemsSource = result;

        //}

    }
}
