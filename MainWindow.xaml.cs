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
            //_excelProgramla = new ExcelProgramla();
            InitializeComponent();
        }

        private void NotVer_Click(object sender, RoutedEventArgs e)
        {
            var excelProgramla = new ExcelProgramla();

            NotVer(excelProgramla);

            void NotVer(ExcelProgramla excel)
            {
                var row = excel.AktifSatir;
                var col = excel.AktifSutun;
                var enAzNot = Convert.ToInt32(txtEnAzNot.Text);

                do
                {
                    var toplaminEsitlenecegiDeger = excel.ReadCellInt(row, excel.ToplaminEsitlenecegiSutun);
                    //excel.ToplamNot = 0;

                    NotlariRasgeleDagit(excel, row, enAzNot, toplaminEsitlenecegiDeger);

                    excel.WriteCell(row, excel.ToplamSutunu, excel.ToplamNot);

                    if (toplaminEsitlenecegiDeger < enAzNot * excel.KriterSayisi)
                    {
                        MessageBox.Show($"Eşitlenecek not en az {enAzNot * excel.KriterSayisi} olmalıdır. ({toplaminEsitlenecegiDeger})");
                        return;
                    }

                    var h = col;

                    if (excel.ToplamNot < toplaminEsitlenecegiDeger)
                    {
                        do
                        {
                            //if (h < excel.ToplamSutunu)
                            //{
                            var artacakSayi = excel.ReadCellInt(row, h);

                            artacakSayi += 5;
                            excel.ToplamNot += 5;

                            excel.WriteCell(row, h, artacakSayi);
                            excel.WriteCell(row, excel.ToplamSutunu, excel.ToplamNot);
                            //}

                            if (h == excel.ToplamSutunu)
                            {
                                h = col;
                            }
                            else
                            {
                                h++;
                            }

                        } while (excel.ToplamNot < toplaminEsitlenecegiDeger);
                    }
                    else
                    {
                        do
                        {
                            var azlacakSayi = excel.ReadCellInt(row, h);

                            if (azlacakSayi > enAzNot)
                            {
                                azlacakSayi -= 5;
                                excel.ToplamNot -= 5;

                                excel.WriteCell(row, h, azlacakSayi);
                                excel.WriteCell(row, excel.ToplamSutunu, excel.ToplamNot);
                            }

                            if (h == excel.ToplamSutunu)
                            {
                                h = col;
                            }
                            else
                            {
                                h++;
                            }

                        } while (excel.ToplamNot > toplaminEsitlenecegiDeger);
                    }

                    row++;

                } while (excel.ReadCellString(row, excel.AktifSutun - 1) != "");

                MessageBox.Show("Notlar başarı ile verildi.");
                //lblSonuc.Content = "Notlar başarı ile verildi.";
            }
        }

        private void NotlariRasgeleDagit(ExcelProgramla excel, int row, int enAzNot, int toplaminEsitlenecegiDeger)
        {
            var r = new Random();

            excel.ToplamNot = 0;

            for (int i = excel.AktifSutun; i < excel.ToplamSutunu; i++)
            {
                if (toplaminEsitlenecegiDeger == 0)
                {
                    excel.ToplamNot = 0;
                    excel.WriteCell(row, i, 0);
                }
                else
                {
                    var a = enAzNot;
                    var b = excel.ReadCellInt(excel.KriterSatiri, i);
                    var hesaplanacakSayi = Math.Round(r.NextDouble() * (b - a) + a, 0);
                    var modHesapla = hesaplanacakSayi % 5;
                    if (modHesapla > 0)
                    {
                        do
                        {
                            hesaplanacakSayi--;
                            modHesapla = hesaplanacakSayi % 5;
                        } while (modHesapla != 0);
                    }
                    excel.WriteCell(row, i, hesaplanacakSayi);
                    excel.ToplamNot += hesaplanacakSayi;
                }
            }
        }

        class ExcelProgramla
        {
            public Workbook WorkBook { get; set; }
            public Worksheet WorkSheet { get; set; }
            public Range Range { get; set; }
            public int AktifSatir { get; set; }
            public int AktifSutun { get; set; }
            public int ToplamSatirSayisi { get; set; }
            public int ToplamSutunSayisi { get; set; }
            public int ToplamSutunu { get; set; }
            public double ToplamNot { get; set; }
            public int ToplaminEsitlenecegiSutun { get; set; }
            public int KriterSatiri => AktifSatir - 1;
            public int KriterSayisi { get; set; }
            public int KriterlerinToplami { get; set; }

            public ExcelProgramla()
            {
                //var excel = new Application();
                var oXL = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                WorkBook = oXL.ActiveWorkbook;
                WorkSheet = WorkBook.ActiveSheet;
                Range = WorkSheet.Application.ActiveCell;
                AktifSatir = Range.Row;
                AktifSutun = Range.Column;
                ParametreleriHesapla();
                //var w = WorkBook.Name;
                //var s = WorkSheet.Name;
            }

            private void ParametreleriHesapla()
            {
                ToplamSatirSayisiHesapla();
                ToplamSutunSayisiHesapla();
                KriterToplamlariHesapla();
            }

            private void KriterToplamlariHesapla()
            {
                for (int i = AktifSutun; i < ToplamSutunu; i++)
                {
                    KriterlerinToplami += ReadCellInt(KriterSatiri, i);
                }

                if (KriterlerinToplami != 100)
                {
                    MessageBox.Show("Kriterlerin toplamı 100 olmalıdır.");
                    return;
                    //throw new Exception("Kriterlerin toplamı 100 olmalıdır.");
                }
            }
            private void ToplamSutunSayisiHesapla()
            {
                ToplamSutunSayisi = AktifSutun;

                do
                {
                    ToplamSutunSayisi++;
                } while (ReadCellString(KriterSatiri, ToplamSutunSayisi) != "");

                ToplaminEsitlenecegiSutun = ToplamSutunSayisi;
                ToplamSutunu = ToplamSutunSayisi - 1;
                ToplamSutunSayisi = ToplamSutunu - 1;
                KriterSayisi = ToplamSutunu - AktifSutun;
            }
            private void ToplamSatirSayisiHesapla()
            {
                ToplamSatirSayisi = AktifSatir;

                while (ReadCellString(ToplamSatirSayisi, AktifSutun - 1) != "")
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

            public int ReadCellInt(int i, int j)
            {
                var deger = WorkSheet.Cells[i, j].Value2;

                if (deger == null)
                {
                    return 0;
                }
                else
                {
                    return Convert.ToInt32(deger);
                }
            }
            public string ReadCellString(int i, int j)
            {
                var deger = WorkSheet.Cells[i, j].Value2;

                if (deger == null)
                {
                    return "";
                }
                else
                {
                    return Convert.ToString(deger);
                }
            }

            public void WriteCell(int i, int j, double deger)
            {
                WorkSheet.Cells[i, j].Value = deger;
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
