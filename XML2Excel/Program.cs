using System;
using System.Data;
using System.Xml;
using System.Data;
using System.IO;
using DocumentFormat.OpenXml;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace XML2Excel
{
    class Program
    {
        
        static void Main(string[] args)
        {

            criaHeader();               
        }

        private static void criaHeader()
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(@"C:\TesteBook.xlsx", SpreadsheetDocumentType.Workbook))
            {
                    //Criar doc, sheet and stuff..
                    spreadSheet.AddWorkbookPart();
                    spreadSheet.WorkbookPart.Workbook = new Workbook ();   
                    spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
                    Worksheet worksheet = spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet = new Worksheet(new SheetData());
                
                    //Ler o XML
                    String xmlString = System.IO.File.ReadAllText(@"C:\Users\User\Desktop\mer.xml");
                    XmlDocument xmlDados = new XmlDocument();
                    xmlDados.LoadXml(xmlString);
                    XmlNamespaceManager manager = new XmlNamespaceManager(xmlDados.NameTable);
                    manager.AddNamespace("my", "http://schemas.microsoft.com/office/infopath/2003/myXSD/2016-03-29T13:18:37");

                    //Para gerar o nome da Sheet
                    String ano = string.Empty;
                    String mapa = string.Empty;
                    String mes = string.Empty;
                    String tipomapa = string.Empty;
                    String validado = string.Empty;
                    foreach (XmlNode node in xmlDados.ChildNodes)
                    {
                        if (node.Name.ToLower() == "my:meuscampos")
                        {
                            ano = node.SelectSingleNode("./my:camposmapa/my:ano", manager).InnerText;
                            mapa = node.SelectSingleNode("./my:camposmapa/my:mapa", manager).InnerText;
                            mes = node.SelectSingleNode("./my:camposmapa/my:mes", manager).InnerText;
                            tipomapa = node.SelectSingleNode("./my:camposmapa/my:postoMapa", manager).InnerText;
                        }
                    }
               
                    //Necessário para fazer merge a células
                    MergeCells mergeCells;
                    MergeCell mergeCell = new MergeCell();
                    if (worksheet.Elements<MergeCells>().Count() > 0)
                    {
                        mergeCells = worksheet.Elements<MergeCells>().First();
                    }
                    else
                    {
                        mergeCells = new MergeCells();

                        if (worksheet.Elements<CustomSheetView>().Count() > 0)
                        {
                            worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                        }
                        else if (worksheet.Elements<DataConsolidate>().Count() > 0)
                        {
                            worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                        }
                        else if (worksheet.Elements<SortState>().Count() > 0)
                        {
                            worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                        }
                        else if (worksheet.Elements<AutoFilter>().Count() > 0)
                        {
                            worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                        }
                        else if (worksheet.Elements<Scenarios>().Count() > 0)
                        {
                            worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                        }
                        else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
                        {
                            worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                        }
                        else if (worksheet.Elements<SheetProtection>().Count() > 0)
                        {
                            worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                        }
                        else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                        {
                            worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                        }
                        else
                        {
                            worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                        }
                    }

                    
                     //Gera-se o conteúdo Linha a linha..
                    //Primeira Linha
                    Row FirstRowHeader = spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.First().AppendChild(new Row());
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue("IDENTIFICAÇÃO"), DataType = new EnumValue<CellValues>(CellValues.String)});
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue("MERGULHADORES (Profundidades em Mts)"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue("GUIAS (Profundidades em Mts)"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue("OBS"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    FirstRowHeader.AppendChild(new Cell() { CellValue = new CellValue("TOTAL"), DataType = new EnumValue<CellValues>(CellValues.String) });

                    //"Unir células"
                    mergeCell = new MergeCell() { Reference = new StringValue("A1" + ":" + "C1") };
                    mergeCells.Append(mergeCell);
                    mergeCell = new MergeCell() { Reference = new StringValue("D1" + ":" + "M1") };
                    mergeCells.Append(mergeCell);
                    mergeCell = new MergeCell() { Reference = new StringValue("N1" + ":" + "W1") };
                    mergeCells.Append(mergeCell);
                    mergeCell = new MergeCell() { Reference = new StringValue("Y1" + ":" + "Z1") };
                    mergeCells.Append(mergeCell);

                    //Segunda linha
                    Row SecondRowHeader = spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.First().AppendChild(new Row());
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("NII"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("POSTO"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("NOME"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("CIRCUITO"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("00_10"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("10_20"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("20_30"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("30_40"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("40_50"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("50_60"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("60_70"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("70_80"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("80_90"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("CIRCUITO"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("00_10"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("10_20"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("20_30"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("30_40"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("40_50"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("50_60"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("60_70"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("70_80"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("80_90"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("M (Min)"), DataType = new EnumValue<CellValues>(CellValues.String) });
                    SecondRowHeader.AppendChild(new Cell() { CellValue = new CellValue("G (Min)"), DataType = new EnumValue<CellValues>(CellValues.String) });

                    //Valores que mudam consoante o XML               
                    criaCorpo(xmlDados, manager, worksheet, mergeCells, mergeCell);

                    //Gravar
                    spreadSheet.WorkbookPart.WorksheetParts.First().Worksheet.Save();

                    //Dar nome a Sheet
                    spreadSheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                    spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet() 
                    { 
                        Id = spreadSheet.WorkbookPart.GetIdOfPart(spreadSheet.WorkbookPart.WorksheetParts.First()), 
                        SheetId = 1,
                        Name = mapa + " (" + tipomapa + ")" + " - " + CultureInfo.CreateSpecificCulture("pt-PT").DateTimeFormat.GetMonthName(Int32.Parse((String)mes)) + " - " + ano
                    });

                    //Gravar e Fechar
                    spreadSheet.WorkbookPart.Workbook.Save();
                    spreadSheet.Close();
            }

        }

        private static void criaCorpo(XmlDocument xmlDados, XmlNamespaceManager manager, Worksheet WS, MergeCells mergeCells, MergeCell mergeCell)
        {
            foreach (XmlNode node in xmlDados.ChildNodes)
            {
                if (node.Name.ToLower() == "my:meuscampos")
                {
                    XmlNodeList mergulhadores = node.SelectNodes("//my:temponii", manager);
                    int j = 0;
                    foreach (XmlNode merg in mergulhadores)
                    {
                        float abono = 0;
                        String keyCA = string.Empty;
                        String keyCF = string.Empty;

                        Row mergulhadorCA = WS.First().AppendChild(new Row());
                        Row mergulhadorCF = WS.First().AppendChild(new Row());
                        String obsTxt = merg.SelectSingleNode("./my:obs", manager).InnerText;
                        String totalMTxt = merg.SelectSingleNode("./my:totalminMer", manager).InnerText;
                        String totalGTxt = merg.SelectSingleNode("./my:totalminGuia", manager).InnerText;
                        String nii = merg.SelectSingleNode("./my:nii", manager).InnerText;
                        String posto = merg.SelectSingleNode("./my:posto", manager).InnerText;
                        String classe = merg.SelectSingleNode("./my:classe", manager).InnerText;
                        String nome = merg.SelectSingleNode("./my:nome", manager).InnerText;

                        String c02Text = merg.SelectSingleNode("./my:M_0_10_A", manager).InnerText;
                        String c12Text = merg.SelectSingleNode("./my:M_0_10_F", manager).InnerText;

                        mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue(nii), DataType = new EnumValue<CellValues>(CellValues.Number) });
                        mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue(posto.ToString()), DataType = new EnumValue<CellValues>(CellValues.String) });
                        mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue(nome.ToString()), DataType = new EnumValue<CellValues>(CellValues.String) });
                        mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue("ABERTO"), DataType = new EnumValue<CellValues>(CellValues.String) });
                        mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue(c02Text), DataType = new EnumValue<CellValues>(CellValues.Number) });
                        mergulhadorCF.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                        mergulhadorCF.AppendChild(new Cell() { CellValue = new CellValue(classe), DataType = new EnumValue<CellValues>(CellValues.String) });
                        mergulhadorCF.AppendChild(new Cell() { CellValue = new CellValue(""), DataType = new EnumValue<CellValues>(CellValues.String) });
                        mergulhadorCF.AppendChild(new Cell() { CellValue = new CellValue("FECHADO"), DataType = new EnumValue<CellValues>(CellValues.String) });
                        mergulhadorCF.AppendChild(new Cell() { CellValue = new CellValue(c12Text), DataType = new EnumValue<CellValues>(CellValues.Number) });

                        for (int i = 1; i < 9; i++)
                        {
                            c02Text = merg.SelectSingleNode("./my:M_" + i + "0_" + (i + 1) + "0_A", manager).InnerText;
                            c12Text = merg.SelectSingleNode("./my:M_" + i + "0_" + (i + 1) + "0_F", manager).InnerText;
                            mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue(c02Text), DataType = new EnumValue<CellValues>(CellValues.Number) });
                            mergulhadorCF.AppendChild(new Cell() { CellValue = new CellValue(c12Text), DataType = new EnumValue<CellValues>(CellValues.Number) });
                        }
                        
                        c02Text = merg.SelectSingleNode("./my:G_0_10_A", manager).InnerText;
                        c12Text = merg.SelectSingleNode("./my:G_0_10_F", manager).InnerText;
                        mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue("ABERTO"), DataType = new EnumValue<CellValues>(CellValues.String) });
                        mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue(c02Text), DataType = new EnumValue<CellValues>(CellValues.Number) });
                        mergulhadorCF.AppendChild(new Cell() { CellValue = new CellValue("FECHADO"), DataType = new EnumValue<CellValues>(CellValues.String) });
                        mergulhadorCF.AppendChild(new Cell() { CellValue = new CellValue(c12Text), DataType = new EnumValue<CellValues>(CellValues.Number) });
                       
                        for (int i = 1; i < 9; i++)
                        {
                            c02Text = merg.SelectSingleNode("./my:G_" + i + "0_" + (i + 1) + "0_A", manager).InnerText;
                            c12Text = merg.SelectSingleNode("./my:G_" + i + "0_" + (i + 1) + "0_F", manager).InnerText;
                            mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue(c02Text), DataType = new EnumValue<CellValues>(CellValues.Number) });
                            mergulhadorCF.AppendChild(new Cell() { CellValue = new CellValue(c12Text), DataType = new EnumValue<CellValues>(CellValues.Number) });
                        }

                        mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue(obsTxt.ToString()), DataType = new EnumValue<CellValues>(CellValues.String) });
                        mergeCell = new MergeCell() { Reference = new StringValue("X" + (j + 3) + ":" + "X" + (j + 4)) };
                        mergeCells.Append(mergeCell);

                        mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue(totalMTxt), DataType = new EnumValue<CellValues>(CellValues.Number) });
                        mergeCell = new MergeCell() { Reference = new StringValue("Y" + (j + 3) + ":" + "Y" + (j + 4)) };
                        mergeCells.Append(mergeCell);

                        mergulhadorCA.AppendChild(new Cell() { CellValue = new CellValue(totalGTxt), DataType = new EnumValue<CellValues>(CellValues.Number) });
                        mergeCell = new MergeCell() { Reference = new StringValue("Z" + (j + 3) + ":" + "Z" + (j + 4)) };
                        mergeCells.Append(mergeCell);
                        j += 2;              
                    }
                }
            }
        }
    }
}
