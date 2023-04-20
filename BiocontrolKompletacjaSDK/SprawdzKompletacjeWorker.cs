using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Soneta.Business;
using Soneta.Handel;
using Soneta.Magazyny;
using Soneta.Towary;
using System;
using System.Drawing;
using System.IO;

//[assembly: AllowPartiallyTrustedCallers]
[assembly: Worker(typeof(BiocontrolKompletacjaSDK.SprawdzKompletacjeWorker), typeof(DokumentHandlowy))]

namespace BiocontrolKompletacjaSDK
{
    internal class SprawdzKompletacjeWorker
    {
        [Context]
        public Context Context { get; set; }

        [Action(
            "Sprawdź poprawność PW",
            Priority = 1000,
            Icon = ActionIcon.Copy,
            Mode = ActionMode.SingleSession,
            Target = ActionTarget.Menu | ActionTarget.ToolbarWithText)]

        public void MyAction()
        {
            INavigatorContext inc = Context[typeof(INavigatorContext)] as INavigatorContext;
            DokumentHandlowy pw = null;
            PozycjaDokHandlowego pw_poz;
            string definicja = "PW - Przyjęcie wewnętrzne";

            object o = inc.SelectedRows[0];

            if (o is DokumentHandlowy)
            {
                pw = (DokumentHandlowy)o;
                if (inc.SelectedRows.Length > 1)
                    throw new Exception("Należy wybrać tylko jeden dokument.");
            }
            else if (o is PozycjaDokHandlowego)
            {
                pw_poz = (PozycjaDokHandlowego)o;
                if (pw_poz.Dokument.Definicja.ToString() == definicja)
                    pw = pw_poz.Dokument;
            }

            if (pw != null && pw.Definicja.ToString() == definicja)
                SprawdzPoprawnosc(pw);
            else
                throw new Exception("Wybrany rekord nie jest dokumentem PW.");
        }

        public void SprawdzPoprawnosc(DokumentHandlowy dokument)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string fileName = dokument.Numer.ToString().Replace("/", "_") + ".xlsx";

            string symbol;
            double ilosc;
            int precyzja, currentRow = 2;
            bool precyzjaPoprawna, stanPoprawny;
            Towar towar;
            View elementy;

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                FilterIndex = 2,
                RestoreDirectory = true,
                InitialDirectory = @"C:\Users\" + Environment.UserName + @"\OneDrive\Dokumenty\",
                FileName = fileName
            };

            bool? result = saveDialog.ShowDialog();

            if (result == true)
            {
                using (var package = new ExcelPackage())
                {
                    string path = @"C:\Users\" + Environment.UserName + @"\OneDrive\Dokumenty\";//saveDialog.FileName;
                    var worksheet = package.Workbook.Worksheets.Add("PW - przyjęcie wewnętrzne");

                    using (Session session = Context.Login.CreateSession(false, false))
                    {
                        HandelModule handelModule = HandelModule.GetInstance(session);
                        TowaryModule towaryModule = TowaryModule.GetInstance(session);
                        MagazynyModule magazynyModule = MagazynyModule.GetInstance(session);

                        foreach (PozycjaDokHandlowego poz in dokument.Pozycje)
                        {
                            // produkt
                            towar = poz.Towar;
                            elementy = towar.ElementyKompletu.CreateView();

                            worksheet.Cells[currentRow, 1].Value = towar.Kod.ToString();
                            worksheet.Cells[currentRow, 2].Value = towar.Nazwa.ToString();
                            worksheet.Cells[currentRow, 4].Value = poz.Ilosc.Value;
                            worksheet.Cells[currentRow, 5].Value = poz.Ilosc.Symbol.ToString();
                            worksheet.Cells[currentRow, 1, currentRow, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[currentRow, 1, currentRow, 10].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                            currentRow++;

                            // elementy kompletu dla produktu
                            foreach (ElementKompletu element in elementy)
                            {
                                if (element.Towar.Nazwa.ToString() != towar.Nazwa.ToString())
                                {
                                    ilosc = element.Ilosc.Value;
                                    symbol = element.Ilosc.Symbol.ToString();

                                    worksheet.Cells[currentRow, 1].Value = element.Towar.Kod.ToString();
                                    worksheet.Cells[currentRow, 2].Value = element.Towar.Nazwa.ToString();
                                    worksheet.Cells[currentRow, 3].Value = ilosc;
                                    worksheet.Cells[currentRow, 4].Value = ilosc * poz.Ilosc.Value;
                                    worksheet.Cells[currentRow, 5].Value = symbol;

                                    // sprawdzenie formatowania jednostki miary
                                    if (ilosc.ToString().Contains(","))
                                        precyzja = ilosc.ToString().Split(',')[1].Length;
                                    else
                                        precyzja = 0;

                                    Jednostka j = towaryModule.Jednostki.WgKodu[symbol];

                                    worksheet.Cells[currentRow, 6].Value = precyzja;
                                    worksheet.Cells[currentRow, 7].Value = j.Precyzja;

                                    precyzjaPoprawna = precyzja <= j.Precyzja;

                                    switch (precyzjaPoprawna)
                                    {
                                        case true:
                                            worksheet.Cells[currentRow, 8].Value = "TAK";
                                            break;
                                        case false:
                                            worksheet.Cells[currentRow, 8].Value = "NIE";
                                            worksheet.Cells[currentRow, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            worksheet.Cells[currentRow, 8].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                            break;
                                    }

                                    // sprawdzenie stanow magazynowych
                                    StanMagazynuWorker smw = new StanMagazynuWorker
                                    {
                                        Towar = element.Towar,
                                        Data = dokument.Data
                                    };

                                    worksheet.Cells[currentRow, 9].Value = smw.Stan.Value;

                                    stanPoprawny = smw.Stan.Value - ilosc * poz.Ilosc.Value >= 0;

                                    switch (stanPoprawny)
                                    {
                                        case true:
                                            worksheet.Cells[currentRow, 10].Value = "TAK";
                                            break;
                                        case false:
                                            worksheet.Cells[currentRow, 10].Value = "NIE";
                                            worksheet.Cells[currentRow, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            worksheet.Cells[currentRow, 10].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                            break;
                                    }

                                    currentRow++;
                                }
                            }
                        }

                        session.Save();
                    }

                    #region headlines
                    worksheet.Cells[1, 1].Value = "Kod towaru";
                    worksheet.Cells[1, 2].Value = "Nazwa towaru";
                    worksheet.Cells[1, 3].Value = "Ilość na produkt";
                    worksheet.Cells[1, 4].Value = "Ilość całkowita";
                    worksheet.Cells[1, 5].Value = "Jednostka";
                    worksheet.Cells[1, 6].Value = "Precyzja jedn. na towarze";
                    worksheet.Cells[1, 7].Value = "Precyzja jedn. w systemie";
                    worksheet.Cells[1, 8].Value = "Precyzja jedn. zgodna";
                    worksheet.Cells[1, 9].Value = "Stan mag.\n " + dokument.Data.ToString();
                    worksheet.Cells[1, 10].Value = "Stan mag. dodatni po operacji";
                    worksheet.Cells[1, 1, 1, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, 1, 1, 10].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    #endregion

                    #region worksheet style
                    for (int row = 1; row <= currentRow - 1; row++)
                        for (int col = 1; col <= 10; col++)
                            worksheet.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    worksheet.Column(8).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Column(10).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Row(1).Style.Font.Bold = true;
                    worksheet.Row(1).Style.WrapText = true;
                    worksheet.Row(1).Height = 50;

                    worksheet.Rows.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.View.FreezePanes(2, 1);
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    for (int col = 4; col <= 10; col++)
                        worksheet.Column(col).Width = 12;
                    #endregion

                    // zapis do pliku
                    package.SaveAs(new FileInfo(path));
                }
            }
        }
    }
}
