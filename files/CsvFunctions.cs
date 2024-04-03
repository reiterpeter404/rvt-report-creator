using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using rvt_report_creator.collector;

namespace rvt_report_creator.files;

public abstract class CsvFunctions
{
     private const string CsvFileExtension = ".csv";
    private const char CsvSeparator = ';';

    /// <summary>
    /// Create the CSV output file.
    /// </summary>
    /// <param name="data">A collection of the data of the export file.</param>
    /// <param name="correctedFilePath">The file path to store the CSV file.</param>
    public static void CreateCsvOutput(List<RvtStatistics?> data, string correctedFilePath)
    {
        StringBuilder stringBuilder = new();

        CreateCsvHeader(stringBuilder);
        CreateCsvSubHeader(stringBuilder);
        foreach (var rvtStatistics in data)
        {
            AppendDataToCsv(stringBuilder, rvtStatistics);
        }

        File.WriteAllText(correctedFilePath + CsvFileExtension, stringBuilder.ToString());
    }

    /// <summary>
    /// Create the header for the CSV file.
    /// </summary>
    /// <param name="stringBuilder">The string builder for the CSV file.</param>
    private static void CreateCsvHeader(StringBuilder stringBuilder)
    {
        bool isFirstElement = true;
        List<string> headerElements = CommonFunctions.LoadHeaderElements();
        foreach (string element in headerElements)
        {
            if (!isFirstElement)
            {
                stringBuilder.Append(CsvSeparator);
            }

            stringBuilder.Append(element);
            isFirstElement = false;
        }

        stringBuilder.AppendLine();
    }

    /// <summary>
    /// Create the sub header for the CSV file.
    /// </summary>
    /// <param name="stringBuilder">The string builder for the CSV file.</param>
    private static void CreateCsvSubHeader(StringBuilder stringBuilder)
    {
        bool isFirstElement = true;
        List<string> subHeaderElements = CommonFunctions.LoadSubHeaderElements();
        foreach (string element in subHeaderElements)
        {
            if (!isFirstElement)
            {
                stringBuilder.Append(CsvSeparator);
            }

            stringBuilder.Append(element);
            isFirstElement = false;
        }

        stringBuilder.AppendLine();
    }

    /// <summary>
    /// Append the data to the CSV file.
    /// </summary>
    /// <param name="stringBuilder">The string builder for the CSV file.</param>
    /// <param name="element">The daily collection of each measurements.</param>
    private static void AppendDataToCsv(StringBuilder stringBuilder, RvtStatistics? element)
    {
        if (element == null)
        {
            return;
        }

        stringBuilder.Append(element.Date.Month);
        stringBuilder.Append(CsvSeparator);
        stringBuilder.Append(element.Date.Day);
        stringBuilder.Append(CsvSeparator); // Tag
        stringBuilder.Append(element.CalculateContainerOutflowPerDay());
        stringBuilder.Append(CsvSeparator);
        stringBuilder.Append(element.CalculateOutFlowPerDay());
        stringBuilder.Append(CsvSeparator);
        stringBuilder.Append(element.CalculatePhMax());
        stringBuilder.Append(CsvSeparator);
        stringBuilder.Append(element.CalculatePhMin());
        stringBuilder.Append(CsvSeparator);
        stringBuilder.Append(element.CalculateTemperatureMean());
        stringBuilder.Append(CsvSeparator);
        stringBuilder.Append(element.CalculateTemperatureMax());


        // append percentiles
        List<TimeSpan> startTime = CommonFunctions.LoadStartTime();
        List<TimeSpan> endTimes = CommonFunctions.LoadEntTime();

        for (int i = 0; i < startTime.Count; i++)
        {
            double temperaturePercentile = element.CreateTemperaturePercentile(startTime[i], endTimes[i], CommonFunctions.DefaultPercentage);
            stringBuilder.Append(CsvSeparator).Append(temperaturePercentile);
        }

        stringBuilder.AppendLine();
    }
}