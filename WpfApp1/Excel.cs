using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    public class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string  path, int sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public void Close()
        {
            wb.Close();
        }

        

        public string[,] ReadRange(int starti, int starty, int endi, int endy) // чтение диапазона
        {
            Range range = (Range) ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value2;
            string[,] returnstring = new string[endi - starti, endy - starty]; 
            for (int i = 1; i <= endi - starti; i++)
            {
                for (int j = 1; j <= endy - starty; j++)
                {
                    returnstring[i - 1, j - 1] = holder[i, j].ToString();
                }
            }
            return returnstring;
        }

        public MyTable ReadRow(int num, int start, int end) // читаем целую строку 
        {
            Range range = (Range)ws.Range[ws.Cells[num, start], ws.Cells[num, end]];
            object[,] holder = range.Value2;
            List<string> returnstring = new List<string>();
            foreach (var i in holder)
            {
                returnstring.Add(i.ToString());
            }
            MyTable table = new MyTable(returnstring[0], returnstring[1], returnstring[2], returnstring[3], returnstring[4], returnstring[5], returnstring[6], returnstring[7], returnstring[8], returnstring[9]);
            return table;
        }

        public string ReadCell(int i, int j) // прочитать ячейку
        {
            i++; j++;
            if (ws.Cells[i, j].Value2 != null)
                return  Convert.ToString(ws.Cells[i, j].Value2);
            else return "blank";
        }
        
    }
}
