using ClosedXML.Excel;
using Microsoft.Win32;
using Soneta.Business;
using Soneta.Handel;
using Soneta.Magazyny;
using Soneta.Towary;
using System;

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
                string path = saveDialog.FileName;

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("PW - przyjęcie wewnętrzne");

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

                            worksheet.Cell(currentRow, 1).Value = towar.Kod.ToString();
                            worksheet.Cell(currentRow, 2).Value = towar.Nazwa.ToString();
                            worksheet.Cell(currentRow, 3).Value = poz.Ilosc.Value;
                            worksheet.Cell(currentRow, 4).Value = poz.Ilosc.Symbol.ToString();
                            worksheet.Range(currentRow, 1, currentRow, 10).Style.Fill.BackgroundColor = XLColor.LightYellow;
                            currentRow++;

                            // elementy kompletu dla produktu
                            foreach (ElementKompletu element in elementy)
                            {
                                if (element.Towar.Nazwa.ToString() != towar.Nazwa.ToString())
                                {
                                    ilosc = element.Ilosc.Value;
                                    symbol = element.Ilosc.Symbol.ToString();

                                    worksheet.Cell(currentRow, 1).Value = element.Towar.Kod.ToString();
                                    worksheet.Cell(currentRow, 2).Value = element.Towar.Nazwa.ToString();
                                    worksheet.Cell(currentRow, 3).Value = ilosc;
                                    worksheet.Cell(currentRow, 4).Value = ilosc * poz.Ilosc.Value;
                                    worksheet.Cell(currentRow, 5).Value = symbol;

                                    // sprawdzenie formatowania jednostki miary
                                    if (ilosc.ToString().Contains(","))
                                        precyzja = ilosc.ToString().Split(',')[1].Length;
                                    else
                                        precyzja = 0;

                                    Jednostka j = towaryModule.Jednostki.WgKodu[symbol];

                                    worksheet.Cell(currentRow, 6).Value = precyzja;
                                    worksheet.Cell(currentRow, 7).Value = j.Precyzja;

                                    precyzjaPoprawna = precyzja <= j.Precyzja;

                                    switch (precyzjaPoprawna)
                                    {
                                        case true:
                                            worksheet.Cell(currentRow, 8).Value = "TAK";
                                            break;
                                        case false:
                                            worksheet.Cell(currentRow, 8).Value = "NIE";
                                            worksheet.Cell(currentRow, 8).Style.Fill.BackgroundColor = XLColor.Red;
                                            break;
                                    }

                                    // sprawdzenie stanow magazynowych
                                    StanMagazynuWorker smw = new StanMagazynuWorker
                                    {
                                        Towar = element.Towar,
                                        Data = dokument.Data
                                    };

                                    worksheet.Cell(currentRow, 9).Value = smw.Stan.Value;

                                    stanPoprawny = smw.Stan.Value - ilosc * poz.Ilosc.Value >= 0;

                                    switch (stanPoprawny)
                                    {
                                        case true:
                                            worksheet.Cell(currentRow, 10).Value = "TAK";
                                            break;
                                        case false:
                                            worksheet.Cell(currentRow, 10).Value = "NIE";
                                            worksheet.Cell(currentRow, 10).Style.Fill.BackgroundColor = XLColor.Red;
                                            break;
                                    }

                                    currentRow++;
                                }
                            }
                        }

                        session.Save();
                    }

                    #region headlines
                    worksheet.Cell(1, 1).Value = "Kod towaru";
                    worksheet.Cell(1, 2).Value = "Nazwa towaru";
                    worksheet.Cell(1, 3).Value = "Ilość na produkt";
                    worksheet.Cell(1, 4).Value = "Ilość całkowita";
                    worksheet.Cell(1, 5).Value = "Jednostka";
                    worksheet.Cell(1, 6).Value = "Precyzja jedn. na towarze";
                    worksheet.Cell(1, 7).Value = "Precyzja jedn. w systemie";
                    worksheet.Cell(1, 8).Value = "Precyzja jedn. zgodna";
                    worksheet.Cell(1, 9).Value = "Stan mag.\n " + dokument.Data.ToString();
                    worksheet.Cell(1, 10).Value = "Stan mag. dodatni po operacji";
                    worksheet.Range(1, 1, 1, 10).Style.Fill.BackgroundColor = XLColor.LightGray;
                    #endregion

                    #region worksheet style
                    for (int row = 1; row <= currentRow - 1; row++)
                        for (int col = 1; col <= 10; col++)
                            worksheet.Cell(row, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

                    worksheet.Column(8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Column(10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Row(1).Style.Font.Bold = true;
                    worksheet.Row(1).Style.Alignment.WrapText = true;
                    worksheet.Row(1).Height = 50;

                    worksheet.Rows().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.SheetView.FreezeRows(1);
                    worksheet.Columns().AdjustToContents();

                    for (int col = 4; col <= 10; col++)
                        worksheet.Column(col).Width = 12;
                    #endregion

                    workbook.SaveAs(path);

                    /*try
                    { 
                        workbook.SaveAs(path); 
                    }
                    catch (Exception ex){}*/
                }               
            }
        }
    }
}
