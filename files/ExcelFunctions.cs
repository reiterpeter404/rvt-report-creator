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
    /// <param name="data">The summary of the exported data.</param>
    /// <param name="filePath">The path of the resulting file, without any file extension.</param>
    public static void CreateExcel(List<RvtStatistics> data, string filePath)
    {
        string excelFilePath = filePath + ExcelExtension;
        using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(excelFilePath, SpreadsheetDocumentType.Workbook);
        WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        Sheets sheets = AppendSummarySheet(data, workbookPart, spreadsheetDocument);

        if (CreateDailyReports)
        {
            AppendDailyReports(data, workbookPart, spreadsheetDocument, sheets);
        }
    }

    /// <summary>
    /// Append the report of several days to the Excel file.
    /// </summary>
    /// <param name="data">The summary of the exported data.</param>
    /// <param name="workbookPart">The workbook part to add the data.</param>
    /// <param name="spreadsheetDocument">The reference to the Excel document</param>
    /// <returns>The sheets containing the sum sheet.</returns>
    private static Sheets AppendSummarySheet(List<RvtStatistics> data, OpenXmlPartContainer workbookPart, SpreadsheetDocument spreadsheetDocument)
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
    /// <param name="sumSheetHeaderRow">The header row to append the data.</param>
    /// <param name="sumSheetSheetData">The sheet data to append the data.</param>
    /// <returns>The current row of the header.</returns>
    private static Row AppendSumSheetHeaders(OpenXmlElement sumSheetHeaderRow, OpenXmlElement? sumSheetSheetData)
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
    /// <param name="data">The summary of the exported data.</param>
    /// <param name="sumSheetSheetData">The sheet data to append the data.</param>
    /// <param name="sumSheetSubHeaderRow">The sub header row to append the data.</param>
    private static void AppendStatisticData(List<RvtStatistics> data, OpenXmlElement sumSheetSheetData, OpenXmlElement sumSheetSubHeaderRow)
    {
        sumSheetSheetData.Append(sumSheetSubHeaderRow);

        foreach (var rvtStatistics in data)
        {
            if (rvtStatistics.Elements.Count < 10)
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

            CalculateAndAppendPercentiles(rvtStatistics, dataRow);
            sumSheetSheetData.Append(dataRow);
        }
    }

    /// <summary>
    /// Calculate and append the percentiles to the row.
    /// </summary>
    /// <param name="rvtStatistics">The collected data of one day.</param>
    /// <param name="dataRow">The current data row.</param>
    private static void CalculateAndAppendPercentiles(RvtStatistics rvtStatistics, OpenXmlElement dataRow)
    {
        List<TimeSpan> startTime = CommonFunctions.LoadStartTime();
        List<TimeSpan> endTimes = CommonFunctions.LoadEntTime();
        for (int i = 0; i < startTime.Count; i++)
        {
            double temperaturePercentile = rvtStatistics.CreateTemperaturePercentile(startTime[i], endTimes[i], CommonFunctions.DefaultPercentage);
            dataRow.Append(
                double.IsNaN(temperaturePercentile)
                    ? new Cell() { CellValue = new CellValue(NanValueString), DataType = CellValues.String }
                    : new Cell() { CellValue = new CellValue(temperaturePercentile), DataType = CellValues.Number }
            );
        }
    }

    /// <summary>
    /// Append a page with the daily report to the Excel file.
    /// </summary>
    /// <param name="data">The summary of the exported data.</param>
    /// <param name="workbookPart">The workbook part to add the data.</param>
    /// <param name="spreadsheetDocument">The reference to the Excel document</param>
    /// <param name="sheets">The sheets containing the sum sheet.</param>
    private static void AppendDailyReports(List<RvtStatistics> data, WorkbookPart workbookPart, SpreadsheetDocument spreadsheetDocument, OpenXmlElement sheets)
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
    /// <param name="worksheetPart">The worksheet part to append the data.</param>
    /// <param name="rvtStatistics">The collected data of one day.</param>
    private static void AppendRvtElementToDailyReport(WorksheetPart worksheetPart, RvtStatistics rvtStatistics)
    {
        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        Row row = new Row();
        row.Append(
            new Cell() { CellValue = new CellValue("Datum und Uhrzeit"), DataType = CellValues.String },
            new Cell() { CellValue = new CellValue("Durchfluss Pufferbehälter [m³/h]"), DataType = CellValues.String },
            new Cell() { CellValue = new CellValue("Durchfluss Mbw. [m³/h]"), DataType = CellValues.String },
            new Cell() { CellValue = new CellValue("Temperatur Mbw. [°C]"), DataType = CellValues.String },
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