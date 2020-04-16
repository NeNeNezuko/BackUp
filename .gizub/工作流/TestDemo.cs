using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Reflection;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Linq;


namespace TestDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            //四字节位置调试
            //float a = 50;
            //byte[] by = BitConverter.GetBytes(a);
            //Console.WriteLine(by[0] + " " + by[1] + " " + by[2] + " " + by[3]);

            //Ticks秒数时间调试
            //DateTime now = new DateTime(637150414590418000);
            //DateTime time_ticks = DateTime.Now.AddDays(3650);
            //long a = time_ticks.Ticks;

            //int SystemTimerCount = (int)(((DateTime.Now.AddHours(-1).Ticks / 10000000) % 1000000000) / 5);
            //Console.WriteLine(SystemTimerCount % 12);


            //小数位调试
            //float a = 1.2345f;
            //float  b = (float)Math.Round(a, 3);
            //string c = "1.23456";  
            //Console.WriteLine(Convert.ToSingle(c));

            //list.where用法
            //List<int> list = new List<int>();
            //List<int> a = new List<int>();
            //list.Add(1);
            //list.Add(2);
            //list.Add(3);
            //list.Add(4);
            //list.Add(5);
            //list.Add(6);
            //list.Add(7);
            //list.Add(8);
            //list.Add(9);
            
            //for (int i=0;i<=list.Count - 1;i++)
            //{
            //    a = list.Where(p => p >= 5).ToList();
            //}
            //for (int i = 0; i <= a.Count - 1; i++)
            //{

            //    Console.WriteLine(a[i]);
            //}


            Console.ReadKey();

            //创建Application对象
            Application xlsApp = new Application(); 
            xlsApp.Visible = true;

            //得到WorkBook对象, 可以用两种方式
            //之一: 打开已有的文件
            Workbook xlsBook = xlsApp.Workbooks.Open(@"D:\Demo\CO零点检查.xls", Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //之二：新建一个文件
            //Workbook xlsBook = xlsApp.Workbooks.Add(Missing.Value);


            //指定要操作的Sheet，两种方式
            //之一：
            Worksheet xlsSheet = (Worksheet)xlsBook.Sheets[1];
            //之二：
            //Excel.Worksheet xlsSheet = (Excel.Worksheet)xlsApp.ActiveSheet;


            //指定单元格，读取数据，两种方法
            //之一：
            Range range1 = xlsSheet.get_Range("C2", Type.Missing);
            if (range1 == null)
            {
                Console.WriteLine(" ");
            }
            else
            {
                Console.WriteLine(range1.Value2);
            }

            //之二：
            Range range2 = (Range)xlsSheet.Cells[2, 3];
            if (range2 == null)
            {
                Console.WriteLine(" ");
            }
            else
            {
                Console.WriteLine(range2.Value2);
            }



            //在单元格中写入数据
            Range range3 = xlsSheet.get_Range("A1", Type.Missing);
            range3.Value2 = "Hello World!";
            range3.Borders.Color = Color.FromArgb(123, 231, 32).ToArgb();
            range3.Font.Color = Color.Red.ToArgb();
            range3.Font.Name = "Arial";
            range3.Font.Size = 9;
            //range3.Orientation = 90;   //vertical
            range3.Columns.HorizontalAlignment = Constants.xlCenter;
            range3.VerticalAlignment = Constants.xlCenter;
            range3.Interior.Color = Color.FromArgb(192, 192, 192).ToArgb();
            range3.Columns.AutoFit();//adjust the column width automatically


            //在某个区域写入数据数组
            int matrixHeight = 20;
            int matrixWidth = 20;
            string[,] martix = new string[matrixHeight, matrixWidth];
            for (int i = 0; i < matrixHeight; i++)
                for (int j = 0; j < matrixWidth; j++)
                {
                    martix[i, j] = String.Format("{0}_{1}", i + 1, j + 1);
                }
            string startColName = GetColumnNameByIndex(0);
            string endColName = GetColumnNameByIndex(matrixWidth - 1);
            //取得某个区域，两种方法
            //之一：
            //Range range4 = xlsSheet.get_Range("A1", Type.Missing);
            //range4 = range4.get_Resize(matrixHeight, matrixWidth);
            ////之二：
            ////Excel.Range range4 = xlsSheet.get_Range(String.Format("{0}{1}", startColName, 1), String.Format("{0}{1}", endColName, martixHeight));
            //range4.Value2 = martix;
            //range4.Font.Color = Color.Red.ToArgb();
            //range4.Font.Name = "Arial";
            //range4.Font.Size = 9;
            //range4.Columns.HorizontalAlignment = Constants.xlCenter;


            //设置column和row的宽度和颜色
            //int columnIndex = 3;
            //int rowIndex = 3;
            //string colName = GetColumnNameByIndex(columnIndex);
            //xlsSheet.get_Range(colName + rowIndex.ToString(), Type.Missing).Columns.ColumnWidth = 20;
            //xlsSheet.get_Range(colName + rowIndex.ToString(), Type.Missing).Rows.RowHeight = 40;
            //xlsSheet.get_Range(colName + rowIndex.ToString(), Type.Missing).Columns.Interior.Color = Color.Blue.ToArgb();//单格颜色
            //xlsSheet.get_Range(5 + ":" + 7, Type.Missing).Rows.Interior.Color = Color.Yellow.ToArgb();//第5行到第7行的颜色
            //                                                                                          //xlsSheet.get_Range("G : G", Type.Missing).Columns.Interior.Color=Color.Pink.ToArgb();//第n列的颜色如何设置？？

            //保存，关闭
            //if (File.Exists(@"D:\Demo\test1.xls"))
            //{
            //    File.Delete(@"D:\Demo\test1.xls");
            //}
            xlsBook.SaveAs(@"D:\Demo\test1.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlsBook.Close(false, Type.Missing, Type.Missing);
            xlsApp.Quit();

            GC.Collect();

            Console.ReadKey();
        }

        //将column index转化为字母，至多两位
        public static string GetColumnNameByIndex(int index)
        {
            string[] alphabet = new string[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            string result = "";
            int temp = index / 26;
            int temp2 = index % 26 + 1;
            if (temp > 0)
            {
                result += alphabet[temp];
            }
            result += alphabet[temp2];
            return result;

        }
    }
}
