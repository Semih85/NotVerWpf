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
                var katsayi = Convert.ToInt32(txtKatsayi.Text);
                var ogrenciSayisi = excel.ToplamSatirSayisi - excel.AktifSatir + 1;
                var notlar = new int[ogrenciSayisi, excel.KriterSayisi + 1];
                var r = new Random();
                var r2 = row;

                if (excel.KriterlerinToplami != 100)
                {
                    MessageBox.Show("Kriterlerin toplamı 100 olmalıdır.");
                    return;
                }
                
                for (int i = 0; i < ogrenciSayisi; i++)
                {
                    var toplaminEsitlenecegiDeger = excel.ReadCellInt(r2, excel.ToplaminEsitlenecegiSutun);
                    var c2 = col;

                    for (int j = 0; j < excel.KriterSayisi; j++)
                    {
                        var a = enAzNot;
                        var b = excel.ReadCellInt(excel.KriterSatiri, c2);
                        var hesaplanacakSayi = Math.Round(r.NextDouble() * (b - a) + a, 0);
                        var modHesapla = hesaplanacakSayi % katsayi;
                        if (modHesapla > 0)
                        {
                            do
                            {
                                hesaplanacakSayi--;
                                modHesapla = hesaplanacakSayi % katsayi;
                            } while (modHesapla != 0);
                        }
                        notlar[i, j] = (int)hesaplanacakSayi;
                        notlar[i, excel.KriterSayisi] += (int)hesaplanacakSayi;//toplam
                        c2++;
                    }

                    if (toplaminEsitlenecegiDeger < enAzNot * excel.KriterSayisi)
                    {
                        MessageBox.Show($"Eşitlenecek not en az {enAzNot * excel.KriterSayisi} olmalıdır. (Girilen Not: {toplaminEsitlenecegiDeger})");
                        return;
                    }

                    var h = 0;
                    var c3 = col;

                    if (notlar[i, excel.KriterSayisi] < toplaminEsitlenecegiDeger)
                    {
                        do
                        {
                            if (notlar[i, h] < excel.ReadCellInt(excel.KriterSatiri, c3))
                            {
                                notlar[i, h] += katsayi;
                                notlar[i, excel.KriterSayisi] += katsayi;
                            }

                            if (h == excel.KriterSayisi - 1)
                            {
                                h = 0;
                                c3 = col;
                            }
                            else
                            {
                                h++;
                                c3++;
                            }

                        } while (notlar[i, excel.KriterSayisi] < toplaminEsitlenecegiDeger);
                    }
                    else
                    {
                        do
                        {
                            if (notlar[i, h] > enAzNot)
                            {
                                notlar[i, h] -= katsayi;
                                notlar[i, excel.KriterSayisi] -= katsayi;
                            }

                            if (h == excel.KriterSayisi - 1)
                            {
                                h = 0;
                            }
                            else
                            {
                                h++;
                            }

                        } while (notlar[i, excel.KriterSayisi] > toplaminEsitlenecegiDeger);
                    }

                    r2++;
                }

                excel.WriteRange(excel.AktifSatir, excel.AktifSutun, excel.ToplamSatirSayisi, excel.ToplamSutunSayisi + 1, notlar);

                //do
                //{
                //    var toplaminEsitlenecegiDeger = excel.ReadCellInt(row, excel.ToplaminEsitlenecegiSutun);
                //    //excel.ToplamNot = 0;

                //    NotlariRasgeleDagit(excel, row, enAzNot, toplaminEsitlenecegiDeger);

                //    excel.WriteCell(row, excel.ToplamSutunu, excel.ToplamNot);

                //    if (toplaminEsitlenecegiDeger < enAzNot * excel.KriterSayisi)
                //    {
                //        MessageBox.Show($"Eşitlenecek not en az {enAzNot * excel.KriterSayisi} olmalıdır. ({toplaminEsitlenecegiDeger})");
                //        return;
                //    }

                //    var h = col;

                //    if (excel.ToplamNot < toplaminEsitlenecegiDeger)
                //    {
                //        do
                //        {
                //            var artacakSayi = excel.ReadCellInt(row, h);

                //            artacakSayi += 5;
                //            excel.ToplamNot += 5;

                //            excel.WriteCell(row, h, artacakSayi);
                //            excel.WriteCell(row, excel.ToplamSutunu, excel.ToplamNot);

                //            if (h == excel.ToplamSutunu)
                //            {
                //                h = col;
                //            }
                //            else
                //            {
                //                h++;
                //            }

                //        } while (excel.ToplamNot < toplaminEsitlenecegiDeger);
                //    }
                //    else
                //    {
                //        do
                //        {
                //            var azlacakSayi = excel.ReadCellInt(row, h);

                //            if (azlacakSayi > enAzNot)
                //            {
                //                azlacakSayi -= 5;
                //                excel.ToplamNot -= 5;

                //                excel.WriteCell(row, h, azlacakSayi);
                //                excel.WriteCell(row, excel.ToplamSutunu, excel.ToplamNot);
                //            }

                //            if (h == excel.ToplamSutunu)
                //            {
                //                h = col;
                //            }
                //            else
                //            {
                //                h++;
                //            }

                //        } while (excel.ToplamNot > toplaminEsitlenecegiDeger);
                //    }

                //    row++;

                //} while (excel.ReadCellString(row, excel.AktifSutun - 1) != "");

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
            public string[,] ReadCellString(int starti, int starty, int endi, int endy)
            {
                var range = (Range)WorkSheet.Range[WorkSheet.Cells[starti, starty], WorkSheet.Cells[endi, endy]];
                object[,] holder = range.Value2;
                string[,] returnString = new string[endi - starti + 1, endy - starty + 1];

                for (int i = 0; i < endi - starti + 1; i++)
                {
                    for (int j = 0; j < endy - starty + 1; j++)
                    {
                        returnString[i, j] = holder[i, j].ToString();
                    }
                }

                return returnString;
            }
            public void WriteRange(int starti, int starty, int endi, int endy, int[,] writeInt)
            {
                var range = (Range)WorkSheet.Range[WorkSheet.Cells[starti, starty], WorkSheet.Cells[endi, endy]];
                range.Value2 = writeInt;
            }

            public void WriteCell(int i, int j, double deger)
            {
                WorkSheet.Cells[i, j].Value2 = deger;
            }

            public void WriteCell(int i, int j, string deger)
            {
                WorkSheet.Cells[i, j].Value2 = deger;
            }

            public void Save()
            {
                WorkBook.Save();
            }

            //public void OpenFile(string path)
            //{
            //    //@"D:\ExcelCSharp\ExcelC\test.xlsx"
            //    var excel = new ExcelProgramla(path, 1);

            //}
        }

    }
}
