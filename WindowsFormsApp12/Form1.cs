using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp12
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;

        private void Form1_Load(object sender, EventArgs e)
        {
            object misValue = System.Reflection.Missing.Value;
            //新建并打开应用
            MyApp = new Excel.Application();
            //新建并打开工作簿
            MyBook = MyApp.Workbooks.Add(misValue);
            //仅打开本地Excel工作簿
            //MyBook = MyApp.Workbooks.Open(DB_PATH);
            MyApp.Visible = true;

            //新建工作表
            //var NewSheet = (Excel.Worksheet)worksheets.Add(worksheets[1],Type.Missing, Type.Missing, Type.Missing);
            //NewSheet.Name = "newsheet";
            //选择工作表
            MySheet = (Excel.Worksheet)MyBook.Worksheets.get_Item(1);

            int rowCount = 9;
            int colCount = 9;

            string[,] data = new string[10,10];


            //TEST CODE
            //data[1, 7] = "HELLO";
            for (int iRow = 1; iRow <= rowCount; iRow++)
            {

                for (int iCol = 1; iCol <= colCount; iCol++)
                {


                    data[iRow, iCol] = "1";
                }

                //data[iRow, 7] = iRow.ToString;
            }



            //MySheet.Cells[1, 1] = data[0, 0];



            for (int iCol = 1; iCol <= colCount; iCol++)
            {
                for (int iRow = 1; iRow <= rowCount; iRow++)
                {

                    MySheet.Cells[iRow][iCol] = data[iRow, iCol];
                }
            }

            //绘制图表
            Excel.ChartObjects xlCharts = (Excel.ChartObjects)MySheet.ChartObjects(Type.Missing);
            Excel.ChartObject xlChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);

            Excel.Chart chartPage = xlChart.Chart;

            chartPage.ChartType = Excel.XlChartType.xlColumnClustered;


            chartPage.SetSourceData(MySheet.get_Range("A2", "B3"), Excel.XlRowCol.xlColumns);

            //导出图片并打开

            chartPage.Export(@"D:\excel_chart_export.bmp", "BMP", misValue);
            pictureBox1.Image = new Bitmap(@"D:\excel_chart_export.bmp");

            //保存并推出
            //MyBook.SaveAs("csharp.net-informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //MyBook.Close(true, misValue, misValue);
            //MyApp.Quit();


        }
    }
}
