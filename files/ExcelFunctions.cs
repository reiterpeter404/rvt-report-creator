// code copied from https://www.c-sharpcorner.com/article/creating-excel-file-using-openxml/

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using rvt_report_creator.collector;

namespace rvt_report_creator.files;

public abstract class ExcelFunctions
{
    private const string NanValueString = "NaN";
    private const string ExcelExtension = ".xlsx";
    private const bool CreateDailyReports = true;

    /// <summary>
    /// Create the Excel file with the given data.
    /// </summary>
    /// <param name="data"></param>
    /// <param name="filePath"></param>
    public static void CreateExcel(List<RvtStatistics> data, string filePath)
    {
        string excelFilePath = filePath + ExcelExtension;
        using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(excelFilePath, SpreadsheetDocumentType.Workbook);
        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var sheets = AppendSummarySheet(data, workbookPart, spreadsheetDocument);

        if(CreateDailyReports)
        {
            AppendDailyReports(data, workbookPart, spreadsheetDocument, sheets);
        }
    }

    /// <summary>
    /// Append the report of several days to the Excel file.
    /// </summary>
    /// <param name="data"></param>
    /// <param name="workbookPart"></param>
    /// <param name="spreadsheetDocument"></param>
    /// <returns></returns>
    private static Sheets AppendSummarySheet(List<RvtStatistics> data, WorkbookPart workbookPart, SpreadsheetDocument spreadsheetDocument)
    {
        // Add sum sheet as the first worksheet
        WorksheetPart sumSheetPart = workbookPart.AddNewPart<WorksheetPart>();
        sumSheetPart.Worksheet = new Worksheet(new SheetData());
        Sheet sumSheetSheet = new Sheet()
        {
            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(sumSheetPart),
            SheetId = 1,
            Name = "Summenblatt"
        };
        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(sumSheetSheet);

        // Add headers to sum sheet
        SheetData sumSheetSheetData = sumSheetPart.Worksheet.GetFirstChild<SheetData>();
        Row sumSheetHeaderRow = new Row();
        var sumSheetSubHeaderRow = AppendSumSheetHeaders(sumSheetHeaderRow, sumSheetSheetData);

        AppendStatisticData(data, sumSheetSheetData, sumSheetSubHeaderRow);

        return sheets;
    }

    /// <summary>
    /// Append the headers to the sum sheet.
    /// </summary>
    /// <param name="sumSheetHeaderRow"></param>
    /// <param name="sumSheetSheetData"></param>
    /// <returns></returns>
    private static Row AppendSumSheetHeaders(Row sumSheetHeaderRow, SheetData? sumSheetSheetData)
    {
        List<string> headerElements = CommonFunctions.LoadHeaderElements();
        foreach (var headerElement in headerElements)
        {
            sumSheetHeaderRow.Append(
                new Cell()
                {
                    CellValue = new CellValue(headerElement),
                    DataType = CellValues.String
                });
        }

        List<string> subHeaderElements = CommonFunctions.LoadSubHeaderElements();
        sumSheetSheetData.Append(sumSheetHeaderRow);

        Row sumSheetSubHeaderRow = new Row();
        foreach (var subHeaderElement in subHeaderElements)
        {
            sumSheetSubHeaderRow.Append(
                new Cell()
                {
                    CellValue = new CellValue(subHeaderElement),
                    DataType = CellValues.String
                });
        }

        return sumSheetSubHeaderRow;
    }

    /// <summary>
    /// Append the statistic data to the sum sheet.
    /// </summary>
    /// <param name="data"></param>
    /// <param name="sumSheetSheetData"></param>
    /// <param name="sumSheetSubHeaderRow"></param>
    private static void AppendStatisticData(List<RvtStatistics> data, SheetData sumSheetSheetData, Row sumSheetSubHeaderRow)
    {
        sumSheetSheetData.Append(sumSheetSubHeaderRow);

        foreach (var rvtStatistics in data)
        {
            if (rvtStatistics?.Elements.Count < 10)
            {
                continue;
            }

            Row dataRow = new Row();
            dataRow.Append(
                new Cell() { CellValue = new CellValue(rvtStatistics.Date.Month), DataType = CellValues.Number },
                new Cell() { CellValue = new CellValue(rvtStatistics.Date.Day), DataType = CellValues.Number },
                new Cell() { CellValue = new CellValue(rvtStatistics.CalculateContainerOutflowPerDay()), DataType = CellValues.Number },
                new Cell() { CellValue = new CellValue(rvtStatistics.CalculateOutFlowPerDay()), DataType = CellValues.Number },
                new Cell() { CellValue = new CellValue(rvtStatistics.CalculatePhMax()), DataType = CellValues.Number },
                new Cell() { CellValue = new CellValue(rvtStatistics.CalculatePhMin()), DataType = CellValues.Number },
                new Cell() { CellValue = new CellValue(rvtStatistics.CalculateTemperatureMean()), DataType = CellValues.Number },
                new Cell() { CellValue = new CellValue(rvtStatistics.CalculateTemperatureMax()), DataType = CellValues.Number }
            );

            // append percentiles
            List<TimeSpan> startTime = CommonFunctions.LoadStartTime();
            List<TimeSpan> endTimes = CommonFunctions.LoadEntTime();
            for (int i = 0; i < startTime.Count; i++)
            {
                double temperaturePercentile = rvtStatistics.CreateTemperaturePercentile(startTime[i], endTimes[i], CommonFunctions.DefaultPercentage);
                if (double.IsNaN(temperaturePercentile))
                {
                    dataRow.Append(
                        new Cell() { CellValue = new CellValue(NanValueString), DataType = CellValues.String }
                    );
                }
                else
                {
                    dataRow.Append(
                        new Cell() { CellValue = new CellValue(temperaturePercentile), DataType = CellValues.Number }
                    );
                }
            }
            sumSheetSheetData.Append(dataRow);
        }
    }

    /// <summary>
    /// Append a page with the daily report to the Excel file.
    /// </summary>
    /// <param name="data"></param>
    /// <param name="workbookPart"></param>
    /// <param name="spreadsheetDocument"></param>
    /// <param name="sheets"></param>
    private static void AppendDailyReports(List<RvtStatistics> data, WorkbookPart workbookPart, SpreadsheetDocument spreadsheetDocument, Sheets sheets)
    {
        foreach (var rvtStatistics in data)
        {
            // skip elements with less than 10 entries
            if (rvtStatistics.Elements.Count < 10)
            {
                continue;
            }

            // Create worksheet for each RvtStatistics
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = (uint)sheets.ChildElements.Count + 1,
                Name = rvtStatistics.Date.ToString("dd-MM")
            };
            sheets.Append(sheet);

            AppendRvtElementToDailyReport(worksheetPart, rvtStatistics);

            worksheetPart.Worksheet.Save();
        }

        workbookPart.Workbook.Save();
    }

    /// <summary>
    /// Append the given RvtStatistics to the daily report.
    /// </summary>
    /// <param name="worksheetPart"></param>
    /// <param name="rvtStatistics"></param>
    private static void AppendRvtElementToDailyReport(WorksheetPart worksheetPart, RvtStatistics rvtStatistics)
    {
        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        Row row = new Row();
        row.Append(
            new Cell() { CellValue = new CellValue("Datum und Uhrzeit"), DataType = CellValues.String },
            new Cell() { CellValue = new CellValue("Durchfluss Pufferbehälter"), DataType = CellValues.String },
            new Cell() { CellValue = new CellValue("Durchfluss Mbw."), DataType = CellValues.String },
            new Cell() { CellValue = new CellValue("Temperatur Mbw."), DataType = CellValues.String },
            new Cell() { CellValue = new CellValue("Ph-Wert Mbw."), DataType = CellValues.String }
        );

        sheetData.Append(row);

        foreach (var rvtStatisticsElement in rvtStatistics.Elements)
        {
            string formattedDateTime = rvtStatisticsElement.DateAndTime.ToString("dd-MM-yyyy HH:mm:ss");

            row = new Row();
            row.Append(
                new Cell() { CellValue = new CellValue(formattedDateTime), DataType = CellValues.String },
                new Cell() { CellValue = new CellValue(rvtStatisticsElement.Pufferbehaelter.Durchfluss), DataType = CellValues.Number },
                new Cell() { CellValue = new CellValue(rvtStatisticsElement.Mbw.Durchfluss), DataType = CellValues.Number },
                new Cell() { CellValue = new CellValue(rvtStatisticsElement.Mbw.Temperature), DataType = CellValues.Number },
                new Cell() { CellValue = new CellValue(rvtStatisticsElement.Mbw.PhWert), DataType = CellValues.Number }
            );
            sheetData.Append(row);
        }
    }
}