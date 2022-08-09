using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;

using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace TestExcelCreate
{

  public class TestModel
  {
    public int TestId { get; set; }
    public string TestName { get; set; }
    public string TestDesc { get; set; }
    public DateTime TestDate { get; set; }
    public bool TestLogic { get; set; }

  }

  public class TestModelList
  {
    public List<TestModel> testData { get; set; }
  }
 
  class Program
  {
    static string fileFullName;
    static SharedStringTablePart sharedStringTablePart = null;

    static public void CreateExcelFile(TestModelList data, string OutPutFileDirectory)
  {
    var datetime = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");

    fileFullName = Path.Combine(OutPutFileDirectory, "Output.xlsx");

    if (File.Exists(fileFullName))
    {
        fileFullName = Path.Combine(OutPutFileDirectory, "Output_" + datetime + ".xlsx");
    }

    SpreadsheetDocument package = SpreadsheetDocument.Create(fileFullName, SpreadsheetDocumentType.Workbook);
//      package.AddCoreFilePropertiesPart();
      CreatePartsForExcel(package, data);
      package.Close();


      validateExcel(fileFullName);
  }

    private static void validateExcel(string fileFullName)
    {
      using (OpenXmlPackage document = SpreadsheetDocument.Open(fileFullName, false))
      {
        var validator = new OpenXmlValidator();
        IEnumerable<ValidationErrorInfo> errors = validator.Validate(document);
        foreach (ValidationErrorInfo info in errors)
        {
          try
          {
            Console.WriteLine("Validation information: {0} {1} in {2} part (path {3}): {4}",
                        info.ErrorType,
                        info.Node.GetType().Name,
                        info.Part.Uri,
                        info.Path.XPath,
                        info.Description);
          }
          catch (Exception ex)
          {
            Console.WriteLine("Validation failed: {0}", ex);
          }
        }
      }
    }

    static private void GenerateWorkbookPartContent(WorkbookPart workbookPart)
    {
      Workbook workbook1 = new Workbook();
      Sheets sheets1 = new Sheets();
      Sheet sheet1 = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };
      sheets1.Append(sheet1);
      workbook1.Append(sheets1);
      workbookPart.Workbook = workbook1;
    }

    static private void CreatePartsForExcel(SpreadsheetDocument document, TestModelList data)
  {

    WorkbookPart workbookPart1 = document.AddWorkbookPart();
      if (workbookPart1.GetPartsOfType<SharedStringTablePart>().Count() > 0)
      {
        sharedStringTablePart = workbookPart1.GetPartsOfType<SharedStringTablePart>().First();
      }
      else
      {
        sharedStringTablePart = workbookPart1.AddNewPart<SharedStringTablePart>();
      }
      SheetData partSheetData = GenerateSheetdataForDetails(data);

      GenerateWorkbookPartContent(workbookPart1);

    WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
      //    GenerateWorkbookStylesPartContent(workbookStylesPart1);
      CreateStylesheet(workbookStylesPart1);
      workbookStylesPart1.Stylesheet.Save();
    WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
      
      GenerateWorksheetPartContent(worksheetPart1, partSheetData);
  }

  static private void GenerateWorksheetPartContent(WorksheetPart worksheetPart, SheetData sheetData)
  {
    Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
    worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
    worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
      worksheet1.AddNamespaceDeclaration("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
      SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1" };

    SheetViews sheetViews1 = new SheetViews();

    SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
    Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };

    sheetView1.Append(selection1);

    sheetViews1.Append(sheetView1);
    SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

    PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
    worksheet1.Append(sheetDimension1);
    worksheet1.Append(sheetViews1);
    worksheet1.Append(sheetFormatProperties1);
    worksheet1.Append(sheetData);
    worksheet1.Append(pageMargins1);
    worksheetPart.Worksheet = worksheet1;
  }


    private static void CreateStylesheet(WorkbookStylesPart workbookStylesPart)
    {
      Stylesheet ss = new Stylesheet();

      Fonts fts = new Fonts();
      DocumentFormat.OpenXml.Spreadsheet.Font ft = new DocumentFormat.OpenXml.Spreadsheet.Font();
      FontName ftn = new FontName();
      ftn.Val = "Calibri";
      FontSize ftsz = new FontSize();
      ftsz.Val = 11;
      ft.FontName = ftn;
      ft.FontSize = ftsz;
      fts.Append(ft);
      fts.Count = (uint)fts.ChildElements.Count;

      Fills fills = new Fills();
      Fill fill;
      PatternFill patternFill;
      fill = new Fill();
      patternFill = new PatternFill();
      patternFill.PatternType = PatternValues.None;
      fill.PatternFill = patternFill;
      fills.Append(fill);
      fill = new Fill();
      patternFill = new PatternFill();
      patternFill.PatternType = PatternValues.Gray125;
      fill.PatternFill = patternFill;
      fills.Append(fill);
      fills.Count = (uint)fills.ChildElements.Count;

      Borders borders = new Borders();
      Border border = new Border();
      border.LeftBorder = new LeftBorder();
      border.RightBorder = new RightBorder();
      border.TopBorder = new TopBorder();
      border.BottomBorder = new BottomBorder();
      border.DiagonalBorder = new DiagonalBorder();
      borders.Append(border);
      borders.Count = (uint)borders.ChildElements.Count;

      CellStyleFormats csfs = new CellStyleFormats();
      CellFormat cf = new CellFormat();
      cf.NumberFormatId = 0;
      cf.FontId = 0;
      cf.FillId = 0;
      cf.BorderId = 0;
      csfs.Append(cf);
      csfs.Count = (uint)csfs.ChildElements.Count;

      uint iExcelIndex = 164;
      NumberingFormats nfs = new NumberingFormats();
      CellFormats cfs = new CellFormats();

      cf = new CellFormat();
      cf.NumberFormatId = 0;
      cf.FontId = 0;
      cf.FillId = 0;
      cf.BorderId = 0;
      cf.FormatId = 0;
      cfs.Append(cf);

      NumberingFormat nf;
      nf = new NumberingFormat();
      nf.NumberFormatId = iExcelIndex++;
      //      nf.FormatCode = "dd/mm/yyyy hh:mm:ss";
      nf.FormatCode = "yyyy-mm-dd hh:mm:ss";
      nfs.Append(nf);
      cf = new CellFormat();
      cf.NumberFormatId = nf.NumberFormatId;
      cf.FontId = 0;
      cf.FillId = 0;
      cf.BorderId = 0;
      cf.FormatId = 0;
      cf.ApplyNumberFormat = true;
      cfs.Append(cf);

      nf = new NumberingFormat();
      nf.NumberFormatId = iExcelIndex++;
      nf.FormatCode = "#,##0.0000";
      nfs.Append(nf);
      cf = new CellFormat();
      cf.NumberFormatId = nf.NumberFormatId;
      cf.FontId = 0;
      cf.FillId = 0;
      cf.BorderId = 0;
      cf.FormatId = 0;
      cf.ApplyNumberFormat = true;
      cfs.Append(cf);

      // #,##0.00 is also Excel style index 4
      nf = new NumberingFormat();
      nf.NumberFormatId = iExcelIndex++;
      nf.FormatCode = "#,##0.00";
      nfs.Append(nf);
      cf = new CellFormat();
      cf.NumberFormatId = nf.NumberFormatId;
      cf.FontId = 0;
      cf.FillId = 0;
      cf.BorderId = 0;
      cf.FormatId = 0;
      cf.ApplyNumberFormat = true;
      cfs.Append(cf);

      // @ is also Excel style index 49
      nf = new NumberingFormat();
      nf.NumberFormatId = iExcelIndex++;
      nf.FormatCode = "@";
      nfs.Append(nf);
      cf = new CellFormat();
      cf.NumberFormatId = nf.NumberFormatId;
      cf.FontId = 0;
      cf.FillId = 0;
      cf.BorderId = 0;
      cf.FormatId = 0;
      cf.ApplyNumberFormat = true;
      cfs.Append(cf);

      nfs.Count = (uint)nfs.ChildElements.Count;
//      cfs.Count = (uint)cfs.ChildElements.Count;

      ss.Append(nfs);
      ss.Append(fts);
      ss.Append(fills);
      ss.Append(borders);
      ss.Append(csfs);
      ss.Append(cfs);

      CellStyles css = new CellStyles();
      CellStyle cs = new CellStyle();
      cs.Name = "Normal";
      cs.FormatId = 0;
      cs.BuiltinId = 0;
      css.Append(cs);
      css.Count = (uint)css.ChildElements.Count;
      ss.Append(css);

      DifferentialFormats dfs = new DifferentialFormats();
      dfs.Count = 0;
      ss.Append(dfs);

      TableStyles tss = new TableStyles();
      tss.Count = 0;
      tss.DefaultTableStyle = "TableStyleMedium9";
      tss.DefaultPivotStyle = "PivotStyleLight16";
      ss.Append(tss);
      workbookStylesPart.Stylesheet = ss;
    }


static private SheetData GenerateSheetdataForDetails(TestModelList data)
  {
    SheetData sheetData1 = new SheetData();
      uint rowIndex = 1;
    sheetData1.Append(CreateHeaderRowForExcel(rowIndex++));

    foreach (TestModel testmodel in data.testData)
    {
      Row partsRows = GenerateRowForChildPartDetail(testmodel, rowIndex++);
      sheetData1.Append(partsRows);
    }
    return sheetData1;
  }

    private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
      // If the part does not contain a SharedStringTable, create one.
      if (shareStringPart.SharedStringTable == null)
      {
        shareStringPart.SharedStringTable = new SharedStringTable();
      }

      int i = 0;

      // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
      foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
      {
        if (item.InnerText == text)
        {
          return i;
        }

        i++;
      }

      // The text does not exist in the part. Create the SharedStringItem and return its index.
      shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
      shareStringPart.SharedStringTable.Save();

      return i;
    }


    static private Row CreateHeaderRowForExcel(uint rowIndex)
    {
    Row workRow = new Row() { RowIndex = rowIndex };
    workRow.Append(CreateCell("Test Id", CellValues.SharedString, rowIndex, 1));
    workRow.Append(CreateCell("Test Name", CellValues.SharedString, rowIndex, 2));
    workRow.Append(CreateCell("Test Date", CellValues.SharedString, rowIndex, 3));
    workRow.Append(CreateCell("Test Logic", CellValues.SharedString, rowIndex, 4));
    return workRow;
  }

    static private string GetExcelColumnName(int columnNumber)
    {
      string columnName = "";

      while (columnNumber > 0)
      {
        int modulo = (columnNumber - 1) % 26;
        columnName = Convert.ToChar('A' + modulo) + columnName;
        columnNumber = (columnNumber - modulo) / 26;
      }

      return columnName;
    }
    //  Below function is used for generating child rows.
    static private Row GenerateRowForChildPartDetail(TestModel testmodel, uint rowIndex)
  {
    Row tRow = new Row();
      var text = testmodel.TestId.ToString();
      tRow.Append(CreateCell(testmodel.TestId, CellValues.Number, rowIndex,  1));
    tRow.Append(CreateCell(testmodel.TestName, CellValues.SharedString, rowIndex, 2));
      text = testmodel.TestDate.ToShortDateString();
      tRow.Append(CreateCell(testmodel.TestDate.ToShortDateString(), CellValues.Date, rowIndex, 3));
     text = testmodel.TestDate.ToShortDateString();

      tRow.Append(CreateCell(testmodel.TestLogic.ToString(), CellValues.Boolean, rowIndex, 4));
    return tRow;
  }
  //Below function is used for creating cell by passing only cell data and it adds default style.
static private Cell CreateCell(Object text, CellValues type, uint rowIndex, int columnIndex)
  {
      var columnName = GetExcelColumnName(columnIndex);
      var cellReference = columnName + rowIndex;

      Cell cell = new Cell() { CellReference = cellReference };
    cell.DataType = type;

      switch (type)
            {
              case CellValues.Date:
          //              DateTime value = DateTime.Today;
          DateTime value = DateTime.Now;
          var ttt = value.ToOADate().ToString();
          var ttt2 = value.ToUniversalTime();
          //cell.CellValue = new CellValue(value.ToString());
          //cell.CellValue = new CellValue(value.ToOADate());
          //          cell.CellValue = new CellValue(value.ToShortDateString());
          cell.CellValue = new CellValue(value.ToUniversalTime());
          cell.StyleIndex = 1U;

          //          cell.StyleIndex = 4;

          break;
              case CellValues.Boolean:
                cell.CellValue = new CellValue(true);
                break;

              case CellValues.Number:
                cell.CellValue = new CellValue((int)text);
                break;

        case CellValues.SharedString:
          int index = InsertSharedStringItem((string)text, sharedStringTablePart);
          cell.CellValue = new CellValue(index.ToString());
          cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

          break;

        default:
                cell.CellValue = new CellValue((string)text);
          break;
            }
      
      //cell.CellValue = new CellValue(text);
      return cell;
  }
static private EnumValue<CellValues> ResolveCellDataTypeOnValue(string text)
  {
    int intVal;
    double doubleVal;
    if (int.TryParse(text, out intVal) || double.TryParse(text, out doubleVal))
    {
      return CellValues.Number;
    }
    else
    {
      return CellValues.String;
    }
  }


    public static void InsertWorksheet(TestModelList data, string docName)
    {
      // Open the document for editing.
      using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
      {
        // Add a blank WorksheetPart.
        WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
        SheetData partSheetData = GenerateSheetdataForDetails(data);
        newWorksheetPart.Worksheet = new Worksheet(partSheetData);

        Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
        string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);

        // Get a unique ID for the new worksheet.
        uint sheetId = 1;
        if (sheets.Elements<Sheet>().Count() > 0)
        {
          sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        }

        // Give the new worksheet a name.
        
        string sheetName = "myNewSheet";

        // Append the new worksheet and associate it with the workbook.
        Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        sheets.Append(sheet);
      }
    }
    // Now, let’s call the main function for generating Excel file into main method with passing our model into it.
  static void Main(string[] args)
  {

    TestModelList tmList = new TestModelList();
    tmList.testData = new List<TestModel>();
    TestModel tm = new TestModel();
    tm.TestId = 1;
    tm.TestName = "Test1";
    tm.TestDesc = "Tested 1 time";
    tm.TestDate = DateTime.Now.Date;
    tm.TestLogic = true;
    tmList.testData.Add(tm);

    TestModel tm1 = new TestModel();
    tm1.TestId = 2;
    tm1.TestName = "Test2";
    tm1.TestDesc = "Tested 2 times";
    tm1.TestDate = DateTime.Now.AddDays(-1);
    tm1.TestLogic = false;
    tmList.testData.Add(tm1);

    TestModel tm2 = new TestModel();
    tm2.TestId = 3;
    tm2.TestName = "Test3";
    tm2.TestDesc = "Tested 3 times";
    tm2.TestDate = DateTime.Now.AddDays(-2);
    tm2.TestLogic = true;

    tmList.testData.Add(tm2);

    TestModel tm3 = new TestModel();
    tm3.TestId = 4;
    tm3.TestName = "Test4";
    tm3.TestDesc = "Tested 4 times";
    tm3.TestDate = DateTime.Now.AddDays(-3);
    tm3.TestLogic = false;
    tmList.testData.Add(tm3);

    CreateExcelFile(tmList, "C:\\TEMP");



      TestModel tm5 = new TestModel();
      tm5.TestId = 5;
      tm5.TestName = "Test5";
      tm5.TestDesc = "Tested 5 times";
      tm5.TestDate = DateTime.Now.AddDays(-8);
      tm5.TestLogic = true;
      tmList.testData.Add(tm5);

 //     InsertWorksheet(tmList, fileFullName);

    Console.ReadLine();

  }
}
}
