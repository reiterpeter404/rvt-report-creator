// code copied from https://www.c-sharpcorner.com/article/creating-excel-file-using-openxml/

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using rvt_report_creator.collector;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace rvt_report_creator.files;

public abstract class ExcelFunctions
{
    private const string NanValueString = "NaN";
    private const string SumSheetName = "Summenblatt";
    private const string SumSheetId = "sumId";
    private const string ContentType = "rId3";
    private const string Prefix = "mc";
    private const string Ignorable = "x14ac";
    private const string Font = "Calibri";
    private const string ExcelExtension = ".xlsx";
    private const uint DefaultStyleIndex = 1U;
    private const string HttpsSchemasOpenXmlFormatsMarkupCompatibility = "https://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string HttpsSchemasMicrosoftComOfficeSpreadsheet = "https://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
    private const string HttpSchemasOpenXmlFormatsOrgOfficeDocumentRelationships = "https://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string HttpSchemasOpenXmlFormatsOrgMarkupCompatibility = "https://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string HttpSchemasMicrosoftComOfficeSpreadSheet = "https://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

    /// <summary>
    /// Create the Excel file with the given data.
    /// </summary>
    /// <param name="data"></param>
    /// <param name="filePath">The file path where the Excel file will be saved.</param>
    public static void CreateExcel(List<RvtStatistics?> data, string filePath)
    {
        string excelFilePath = filePath + ExcelExtension;
        using SpreadsheetDocument document = SpreadsheetDocument.Create(excelFilePath, SpreadsheetDocumentType.Workbook);
        SheetData mainSheet = GenerateSumSheet(data);
        WorkbookPart workbook = GenerateWorkbookContent(document);
        GenerateWorkbookStylesContent(workbook);
        GenerateWorksheetContent(workbook, mainSheet);
    }

    /// <summary>
    /// Add the cells to the main sheet.
    /// </summary>
    /// <param name="data">The data containing the values.</param>
    /// <returns></returns>
    private static SheetData GenerateSumSheet(List<RvtStatistics?> data)
    {
        List<string> headerElements = CommonFunctions.LoadHeaderElements();
        List<string> subHeaderElements = CommonFunctions.LoadSubHeaderElements();
        SheetData sheetData = new SheetData();
        sheetData.Append(CreateHeaderRow(headerElements));
        sheetData.Append(CreateHeaderRow(subHeaderElements));
        sheetData.Append(CreateEmptyRow(headerElements.Count));

        foreach (RvtStatistics? element in data)
        {
            GenerateRowForChildPartDetail(element, sheetData);
        }

        return sheetData;
    }

    /// <summary>
    /// Append the statistics for each day of the given element to the sheet.
    /// </summary>
    /// <param name="element">The object that holds the elements of a day.</param>
    /// <param name="sheetData">The sheet data that will be added to the Excel file.</param>
    private static void GenerateRowForChildPartDetail(RvtStatistics? element, OpenXmlElement sheetData)
    {
        // avoid elements with less than 10 entries
        if (element?.Elements.Count < 10)
        {
            return;
        }

        Row tRow = new Row();
        tRow.Append(CreateNumberCell(element.Date.Month));
        tRow.Append(CreateNumberCell(element.Date.Day));
        tRow.Append(CreateDecimalCell(element.CalculateContainerOutflowPerDay()));
        tRow.Append(CreateDecimalCell(element.CalculateOutFlowPerDay()));
        tRow.Append(CreateDecimalCell(element.CalculatePhMax()));
        tRow.Append(CreateDecimalCell(element.CalculatePhMin()));
        tRow.Append(CreateDecimalCell(element.CalculateTemperatureMean()));
        tRow.Append(CreateDecimalCell(element.CalculateTemperatureMax()));

        // append percentiles
        List<TimeSpan> startTime = CommonFunctions.LoadStartTime();
        List<TimeSpan> endTimes = CommonFunctions.LoadEntTime();

        for (int i = 0; i < startTime.Count; i++)
        {
            double temperaturePercentile = element.CreateTemperaturePercentile(startTime[i], endTimes[i], CommonFunctions.DefaultPercentage);
            tRow.Append(CreateDecimalCell(temperaturePercentile));
        }

        sheetData.Append(tRow);
    }

    /// <summary>
    /// Create the header row for the Excel file.
    /// </summary>
    /// <param name="elements"></param>
    /// <returns></returns>
    private static Row CreateHeaderRow(List<string> elements)
    {
        Row workRow = new Row();
        foreach (string element in elements)
        {
            workRow.Append(CreateTextCell(element, 2U));
        }

        return workRow;
    }

    /// <summary>
    /// Create an empty row with the given count of cells.
    /// </summary>
    /// <param name="count"></param>
    /// <returns></returns>
    private static Row CreateEmptyRow(int count)
    {
        Row workRow = new Row();
        for (int i = 0; i < count; i++)
        {
            workRow.Append(CreateTextCell("", 2U));
        }

        return workRow;
    }

    /// <summary>
    /// Generate the workbook content for the Excel file.
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    private static WorkbookPart GenerateWorkbookContent(SpreadsheetDocument document)
    {
        Sheets sheets = new();
        sheets.Append(new Sheet()
        {
            Name = SumSheetName,
            SheetId = (UInt32Value)DefaultStyleIndex,
            Id = SumSheetId
        });

        Workbook workbook = new();
        workbook.Append(sheets);

        WorkbookPart workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = workbook;
        return workbookPart;
    }

    /// <summary>
    /// Generate the content for the Excel file.
    /// </summary>
    /// <param name="workbookPart"></param>
    /// <param name="sheetData"></param>
    private static void GenerateWorksheetContent(OpenXmlPartContainer workbookPart, OpenXmlElement sheetData)
    {
        Worksheet worksheet = new Worksheet()
        {
            MCAttributes = new MarkupCompatibilityAttributes()
            {
                Ignorable = Ignorable
            }
        };
        worksheet.AddNamespaceDeclaration("r", HttpSchemasOpenXmlFormatsOrgOfficeDocumentRelationships);
        worksheet.AddNamespaceDeclaration(Prefix, HttpSchemasOpenXmlFormatsOrgMarkupCompatibility);
        worksheet.AddNamespaceDeclaration(Ignorable, HttpSchemasMicrosoftComOfficeSpreadSheet);

        OpenXmlElement[] worksheetChildren = new OpenXmlElement[]
        {
            new SheetDimension() { Reference = "A1" },
            CreateSheetViews(),
            CreateSheetFormatProperties(),
            sheetData,
            CreatePageMargins()
        };

        worksheet.Append(worksheetChildren);

        WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>(SumSheetId);
        worksheetPart.Worksheet = worksheet;
    }

    private static SheetViews CreateSheetViews()
    {
        Selection selection = new Selection()
        {
            ActiveCell = "A1",
            SequenceOfReferences = new ListValue<StringValue>()
            {
                InnerText = "A1"
            }
        };

        SheetView sheetView = new SheetView()
        {
            TabSelected = true,
            WorkbookViewId = (UInt32Value)0U
        };
        sheetView.Append(selection);

        SheetViews sheetViews = new SheetViews();
        sheetViews.Append(sheetView);
        return sheetViews;
    }

    private static SheetFormatProperties CreateSheetFormatProperties()
    {
        SheetFormatProperties sheetFormatProperties = new SheetFormatProperties()
        {
            DefaultRowHeight = 15D,
            DyDescent = 0.25D
        };
        return sheetFormatProperties;
    }

    private static PageMargins CreatePageMargins()
    {
        PageMargins pageMargins = new PageMargins()
        {
            Left = 0.7D,
            Right = 0.7D,
            Top = 0.75D,
            Bottom = 0.75D,
            Header = 0.3D,
            Footer = 0.3D
        };
        return pageMargins;
    }

    /// <summary>
    /// Generate the styles for the Excel file.
    /// </summary>
    /// <param name="workbookPart1"></param>
    private static void GenerateWorkbookStylesContent(OpenXmlPartContainer workbookPart1)
    {
        Stylesheet stylesheet = new Stylesheet()
        {
            MCAttributes = new MarkupCompatibilityAttributes()
            {
                Ignorable = Ignorable
            }
        };
        stylesheet.AddNamespaceDeclaration(Prefix, HttpsSchemasOpenXmlFormatsMarkupCompatibility);
        stylesheet.AddNamespaceDeclaration(Ignorable, HttpsSchemasMicrosoftComOfficeSpreadsheet);

        OpenXmlElement[] styleSheetElements = new OpenXmlElement[]
        {
            CreateFonts(),
            CreateFills(),
            CreateBorders(),
            CreateCellStyleFormats(),
            CreateCellFormats(),
            CreateCellStyles(),
            CreateDifferentialFormats(),
            CreateTableStyles(),
            CreateStylesheetExtensions()
        };

        stylesheet.Append(styleSheetElements);

        WorkbookStylesPart workbookStylesPart = workbookPart1.AddNewPart<WorkbookStylesPart>(ContentType);
        workbookStylesPart.Stylesheet = stylesheet;
    }

    /// <summary>
    /// Create the fonts for the Excel file.
    /// </summary>
    /// <returns></returns>
    private static Fonts CreateFonts()
    {
        Fonts fonts = new Fonts() { Count = (UInt32Value)2U, KnownFonts = true };
        AppendFont(fonts, false);
        AppendFont(fonts, true);
        return fonts;
    }

    /// <summary>
    /// Append a font to the Excel file.
    /// </summary>
    /// <param name="fonts"></param>
    /// <param name="bold"></param>
    private static void AppendFont(OpenXmlElement fonts, bool bold)
    {
        Font font = new Font();
        if (bold)
        {
            font.Append(new Bold());
        }

        font.Append(new FontSize() { Val = 11D });
        font.Append(new Color() { Theme = (UInt32Value)DefaultStyleIndex });
        font.Append(new FontName() { Val = Font });
        font.Append(new FontFamilyNumbering() { Val = 2 });
        font.Append(new FontScheme() { Val = FontSchemeValues.Minor });
        fonts.Append(font);
    }

    /// <summary>
    /// Create the fills for the Excel file.
    /// </summary>
    /// <returns></returns>
    private static Fills CreateFills()
    {
        Fills fills1 = new Fills() { Count = (UInt32Value)2U };
        Fill fill1 = new Fill();
        PatternFill patternFill1 = new PatternFill()
        {
            PatternType = PatternValues.None
        };

        fill1.Append(patternFill1);

        Fill fill2 = new Fill();
        PatternFill patternFill2 = new PatternFill()
        {
            PatternType = PatternValues.Gray125
        };

        fill2.Append(patternFill2);

        fills1.Append(fill1);
        fills1.Append(fill2);
        return fills1;
    }

    /// <summary>
    /// Create the borders for the Excel file.
    /// </summary>
    /// <returns></returns>
    private static Borders CreateBorders()
    {
        Borders borders = new Borders()
        {
            Count = (UInt32Value)2U
        };

        CreateBorders(borders);
        CreateCustomBorders(borders);
        return borders;
    }

    private static void CreateBorders(OpenXmlElement borders)
    {
        Border border = new Border();
        border.Append(new LeftBorder());
        border.Append(new RightBorder());
        border.Append(new TopBorder());
        border.Append(new BottomBorder());
        border.Append(new DiagonalBorder());
        borders.Append(border);
    }

    private static void CreateCustomBorders(OpenXmlElement borders)
    {
        LeftBorder leftBorder = new LeftBorder() { Style = BorderStyleValues.Thin };
        leftBorder.Append(CreateColorElement());

        RightBorder rightBorder = new RightBorder() { Style = BorderStyleValues.Thin };
        rightBorder.Append(CreateColorElement());

        TopBorder topBorder = new TopBorder() { Style = BorderStyleValues.Thin };
        topBorder.Append(CreateColorElement());

        BottomBorder bottomBorder = new BottomBorder() { Style = BorderStyleValues.Thin };
        bottomBorder.Append(CreateColorElement());

        DiagonalBorder diagonalBorder = new DiagonalBorder();

        Border border = new Border();
        border.Append(leftBorder);
        border.Append(rightBorder);
        border.Append(topBorder);
        border.Append(bottomBorder);
        border.Append(diagonalBorder);

        borders.Append(border);
    }

    private static Color CreateColorElement()
    {
        return new Color()
        {
            Indexed = (UInt32Value)64U
        };
    }

    /// <summary>
    /// Create the cell style formats for the Excel file.
    /// </summary>
    /// <returns></returns>
    private static CellStyleFormats CreateCellStyleFormats()
    {
        CellStyleFormats cellStyleFormats = new CellStyleFormats() { Count = (UInt32Value)DefaultStyleIndex };
        CellFormat cellFormat = new CellFormat()
        {
            NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U,
            BorderId = (UInt32Value)0U
        };

        cellStyleFormats.Append(cellFormat);
        return cellStyleFormats;
    }

    /// <summary>
    /// Create the cell formats for the Excel file.
    /// </summary>
    /// <returns></returns>
    private static CellFormats CreateCellFormats()
    {
        CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)3U };
        AppendCellFormat(cellFormats);
        AppendCellFormatUsingBorder(cellFormats);
        AppendCellFormatUsingFontAndBorder(cellFormats);
        return cellFormats;
    }

    private static void AppendCellFormat(OpenXmlElement cellFormats)
    {
        CellFormat cellFormat = new CellFormat()
        {
            NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U,
            BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U
        };
        cellFormats.Append(cellFormat);
    }

    private static void AppendCellFormatUsingBorder(OpenXmlElement cellFormats)
    {
        CellFormat cellFormat = new CellFormat()
        {
            NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U,
            BorderId = (UInt32Value)DefaultStyleIndex, FormatId = (UInt32Value)0U, ApplyBorder = true
        };
        cellFormats.Append(cellFormat);
    }

    private static void AppendCellFormatUsingFontAndBorder(OpenXmlElement cellFormats)
    {
        CellFormat cellFormat = new CellFormat()
        {
            NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)DefaultStyleIndex, FillId = (UInt32Value)0U,
            BorderId = (UInt32Value)DefaultStyleIndex, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true
        };
        cellFormats.Append(cellFormat);
    }

    /// <summary>
    /// Create the cell styles for the Excel file.
    /// </summary>
    /// <returns></returns>
    private static CellStyles CreateCellStyles()
    {
        CellStyle cellStyle = new CellStyle()
        {
            Name = "Normal",
            FormatId = (UInt32Value)0U,
            BuiltinId = (UInt32Value)0U
        };

        CellStyles cellStyles = new CellStyles() { Count = (UInt32Value)DefaultStyleIndex };
        cellStyles.Append(cellStyle);
        return cellStyles;
    }

    /// <summary>
    /// Create the differential formats for the Excel file.
    /// </summary>
    /// <returns></returns>
    private static DifferentialFormats CreateDifferentialFormats()
    {
        return new DifferentialFormats() { Count = (UInt32Value)0U };
    }

    /// <summary>
    /// Create the table styles for the Excel file.
    /// </summary>
    /// <returns></returns>
    private static TableStyles CreateTableStyles()
    {
        return new TableStyles()
        {
            Count = (UInt32Value)0U,
            DefaultTableStyle = "TableStyleMedium2",
            DefaultPivotStyle = "PivotStyleLight16"
        };
    }

    /// <summary>
    /// Create the stylesheet extensions for the Excel file.
    /// </summary>
    /// <returns></returns>
    private static StylesheetExtensionList CreateStylesheetExtensions()
    {
        StylesheetExtensionList stylesheetExtensionList = new StylesheetExtensionList();

        X14.SlicerStyles slicerStyles = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };
        {
            StylesheetExtension stylesheetExtension = new StylesheetExtension()
            {
                Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"
            };
            stylesheetExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            stylesheetExtension.Append(slicerStyles);
            stylesheetExtensionList.Append(stylesheetExtension);
        }
        {
            X15.TimelineStyles timelineStyles = new X15.TimelineStyles()
            {
                DefaultTimelineStyle = "TimeSlicerStyleLight1"
            };

            StylesheetExtension stylesheetExtension = new StylesheetExtension()
            {
                Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}"
            };
            stylesheetExtension.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            stylesheetExtension.Append(timelineStyles);
            stylesheetExtensionList.Append(stylesheetExtension);
        }

        return stylesheetExtensionList;
    }

    /// <summary>
    /// Create a new cell with the given text.
    /// </summary>
    /// <param name="text"></param>
    /// <param name="styleIndex"></param>
    /// <returns></returns>
    private static Cell CreateTextCell(string text, uint styleIndex = DefaultStyleIndex)
    {
        Cell cell = new()
        {
            StyleIndex = styleIndex,
            DataType = CellValues.String,
            CellValue = new CellValue(text)
        };
        return cell;
    }

    /// <summary>
    /// Create a new cell with the given number.
    /// </summary>
    /// <param name="number"></param>
    /// <param name="styleIndex"></param>
    /// <returns></returns>
    private static Cell CreateNumberCell(int number, uint styleIndex = DefaultStyleIndex)
    {
        Cell cell = new()
        {
            StyleIndex = styleIndex,
            DataType = CellValues.Number,
            CellValue = new CellValue(number)
        };
        return cell;
    }

    /// <summary>
    /// Create a new cell with the given number. If the number is NaN, the cell will contain "NaN".
    /// </summary>
    /// <param name="number">The number to insert to the cell.</param>
    /// <param name="styleIndex">The style of the cell.</param>
    /// <returns>The cell object.</returns>
    private static Cell CreateDecimalCell(double number, uint styleIndex = DefaultStyleIndex)
    {
        if (double.IsNaN(number))
        {
            return new Cell()
            {
                StyleIndex = styleIndex,
                DataType = CellValues.String,
                CellValue = new CellValue(NanValueString)
            };
        }

        return new Cell()
        {
            StyleIndex = styleIndex,
            DataType = CellValues.Number,
            CellValue = new CellValue(number)
        };
    }
}