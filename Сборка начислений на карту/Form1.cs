using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Сборка_начислений_на_карту
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            if (DateTime.Today.Year != 2022 )
            {
                //MessageBox.Show(Convert.ToString(DateTime.Today.Year));
                this.Close();
            }

            
        }
        object[,] srcArr_ob;//1 - строка, 2 - столбец

         void button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
             ArrayList itogo = new ArrayList();//маассив для итоговых значений четные = имена, нечетные = значения.
             itogo.Add("");
            
             if (folderBrowserDialog1.SelectedPath.ToString()!="")
             {
                 //MessageBox.Show(folderBrowserDialog1.SelectedPath.ToString());
                 string[] spisokF = System.IO.Directory.GetFiles(folderBrowserDialog1.SelectedPath.ToString());
                 progressBar1.Value = 0;
                 progressBar1.Maximum = spisokF.Length;
                 button1.Text = "Подготовка...";
                 button1.Refresh();
                 Excel.Application APExcel = new Microsoft.Office.Interop.Excel.Application();
                 APExcel.DisplayAlerts = false;
                 APExcel.Visible = false;

                 for (int i = 0; i < spisokF.Length; i++)
                 {
                     srcArr_ob = null;
                     APExcel.Workbooks.Open(spisokF[i], Type.Missing, false, Type.Missing, Type.Missing, Type.Missing,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
    Type.Missing, Type.Missing);
                     Excel.Worksheet Excelsheet_ob = (Excel.Worksheet)APExcel.Sheets[1]; //определяем рабочий лист
                     srcArr_ob = (object[,])Excelsheet_ob.UsedRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);//забираем весь лист
                     APExcel.Workbooks.Close();//закрываем книгу
                     //Формируем массив 
                    //MessageBox.Show(srcArr_ob[12, 4].ToString());

                     for (int q = 1; q < srcArr_ob.GetUpperBound(0);q++ )
                     {
                         if (itogo.Contains(srcArr_ob[q, 5]))
                         {
                             itogo[itogo.LastIndexOf(srcArr_ob[q, 5]) + 1] = Convert.ToString(Convert.ToDouble(itogo[itogo.LastIndexOf(srcArr_ob[q, 5].ToString()) + 1]) + Convert.ToDouble(srcArr_ob[q, 9]));
                         }
                         else
                         {
                             if (Convert.ToString(srcArr_ob[q, 5]) != "")
                             {
                                 if (Convert.ToString(srcArr_ob[q, 9]) != "")
                                 {
                                     if (Convert.ToString(srcArr_ob[q, 5]).Length > 3)
                                     {
                                         if (!Convert.ToString(srcArr_ob[q, 5]).Contains(","))
                                         { 
                                         itogo.Add(srcArr_ob[q, 5]);
                                         itogo.Add(srcArr_ob[q, 9]);
                                         //MessageBox.Show(Convert.ToString(srcArr_ob[q, 4]));
                                         }
                                     }
                                 }
                             }
                         }
                         //itogo.Capacity
                         
                         

                     }

                     //MessageBox.Show(srcArr_ob.Length.ToString());
                     progressBar1.Value = i+1;
                     button1.Text = "Сбор файлов " + progressBar1.Value.ToString() + " / " + spisokF.Length.ToString();
                     button1.Refresh();
                     progressBar1.Refresh();
                     //MessageBox.Show(itogo[itogo.Count-1].ToString());


                 }
                 //формируем списом файл отчета
                 progressBar1.Value = 0;
                 progressBar1.Maximum = itogo.Count;
                 button1.Text = "Формируем отчет...";
                 button1.Refresh();
                 progressBar1.Refresh();

                 Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
                 Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
                 //Книга.
                 ExcelWorkBook = APExcel.Workbooks.Add(System.Reflection.Missing.Value);
                 //Таблица.
                 ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                 int r = 1;
                 for (int w=1;w<itogo.Count;w++)
                 {
                    
                     APExcel.Cells[r,1] = itogo[w];
                     APExcel.Cells[r, 2] = itogo[w+1];
                     w++;
                     r++;
                     progressBar1.Value = w;
                     progressBar1.Refresh();
                 }

                 // Форматируем таблицу
                 ExcelWorkSheet.Columns[2, Type.Missing].NumberFormat = "# ###,##";
                 ExcelWorkSheet.Columns[2, Type.Missing].Style = "Currency";
                 ExcelWorkSheet.Columns[1, Type.Missing].EntireColumn.AutoFit();
                 ExcelWorkSheet.Columns[2, Type.Missing].EntireColumn.AutoFit();

                 //Завершаем 
                 button1.Text = "Готово";
                 button1.Enabled = false;
                 progressBar1.Value = 0;
                 APExcel.DisplayAlerts = true;
                 APExcel.Visible = true;
                 //конец области
             }
            
        }

         private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
         {
             Form ab1 = new AboutBox1();
             ab1.Show();
         }

         private void настроToolStripMenuItem_Click(object sender, EventArgs e)
         {
             Form F2 = new Form2();
             F2.ShowDialog();
         }
    }
}
