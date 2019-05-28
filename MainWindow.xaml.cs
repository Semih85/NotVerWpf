using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
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

namespace NotVerWpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void NotVer_Click(object sender, RoutedEventArgs e)
        {
            var excel2 = new ExcelProgramla();

            NotVer(excel2, 4, 10);

            void NotVer(ExcelProgramla excel, int sutunNumarasi, int toplamSutunSayisi)
            {
                var row = excel.RowNumber;
                var col = excel.ColumnNumber;

                //int[,] rr = new int[24, 7];
                do
                {
                    var t = 0;
                    var h = 0;
                    //Console.WriteLine("--- r: " + row);
                    for (int j = sutunNumarasi; j <= toplamSutunSayisi; j++)
                    {
                        //var deger = excel.ReadCell<int>(row, col);
                        var a = 5;

                        var b = excel.ReadCell<int>(row - (1 + t), h);
                        excel.WriteCell(row, col, 5);
                        //Console.WriteLine(deger);
                        col++;
                    }
                    row++;
                    col = excel.ColumnNumber;

                } while (excel.ReadCell<string>(row, sutunNumarasi - 1) != null);

            }
        }

        class ExcelProgramla
        {
            public Workbook WorkBook { get; set; }
            public Worksheet WorkSheet { get; set; }
            public Range Range { get; set; }
            public int RowNumber { get; set; }
            public int ColumnNumber { get; set; }
            public int ToplamSatirSayisi { get; set; }
            public int ToplamSutunSayisi { get; set; }
            public int ToplamSutunu { get; set; }
            public int ToplaminEsitlenecegiSutun { get; set; }
            public int KriterSatiri => RowNumber - 1;
            public int KriterSayisi { get; set; }
            public int KriterlerinToplami { get; set; }

            public ExcelProgramla()
            {
                //var excel = new Application();
                var oXL = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                WorkBook = oXL.ActiveWorkbook;
                WorkSheet = WorkBook.ActiveSheet;
                Range = WorkSheet.Application.ActiveCell;
                RowNumber = Range.Row;
                ColumnNumber = Range.Column;
                ToplamSatirSayisiHesapla();
                ToplamSutunSayisiHesapla();
                KriterToplamlariHesapla();
                //var w = WorkBook.Name;
                //var s = WorkSheet.Name;
            }

            private void ToplamSutunSayisiHesapla()
            {
                ToplamSutunSayisi = ColumnNumber;

                do
                {
                    ToplamSutunSayisi++;
                } while (ReadCell<int>(KriterSatiri, ToplamSutunSayisi) != 0);

                ToplaminEsitlenecegiSutun = ToplamSutunSayisi;
                ToplamSutunu = ToplamSutunSayisi - 1;
                ToplamSutunSayisi = ToplamSutunu - 1;
                KriterSayisi = ToplamSutunu - ColumnNumber;
            }

            private void KriterToplamlariHesapla()
            {
                for (int i = ColumnNumber; i < ToplamSutunu; i++)
                {
                    KriterlerinToplami += ReadCell<int>(KriterSatiri, i);
                }

                if (KriterlerinToplami != 100)
                {
                    throw new Exception("Kriterlerin toplamı 100 olmalıdır.");
                }
            }

            private void ToplamSatirSayisiHesapla()
            {
                ToplamSatirSayisi = RowNumber;

                while (ReadCell<string>(ToplamSatirSayisi, ColumnNumber - 1) != null && ReadCell<string>(ToplamSatirSayisi, ColumnNumber - 1) != "")
                {
                    ToplamSatirSayisi++;
                }

                ToplamSatirSayisi--;
            }

            public ExcelProgramla(string path, int sheet)
            {
                var excel = new Microsoft.Office.Interop.Excel.Application();

                WorkBook = excel.Workbooks.Open(path);
                WorkSheet = WorkBook.Worksheets[sheet];
            }

            public T ReadCell<T>(int i, int j)
            {
                return (T)WorkSheet.Cells[i, j].Value2;
            }

            public void WriteCell(int i, int j, double deger)
            {
                WorkSheet.Cells[i, j].Value2 = deger;
            }

            public void WriteCell(int i, int j, string deger)
            {
                WorkSheet.Cells[i, j].Value2 = deger;
            }

            //public void OpenFile(string path)
            //{
            //    //@"D:\ExcelCSharp\ExcelC\test.xlsx"
            //    var excel = new ExcelProgramla(path, 1);

            //}
        }
    }
}
