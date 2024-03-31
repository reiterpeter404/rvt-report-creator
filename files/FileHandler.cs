using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using rvt_report_creator.collector;
using rvt_report_creator.measurements;

namespace rvt_report_creator.files;

/// <summary>
/// The FileHandler class is used to read and write text files.
/// </summary>
public abstract class FileHandler
{
    private const bool CreateExcelOutput = true;
    private const string DateAndTimeString = "Datum und Uhrzeit;";

    /// <summary>
    /// Read the file and prepare the data for further processing.
    /// </summary>
    /// <param name="filePath">The file path of the file.</param>
    /// <returns>A matrix of the data in the file.</returns>
    public static List<RvtElement> ReadFile(string filePath)
    {
        string fileContent = ReadTextFile(filePath);
        List<string> lines = fileContent.Split("\r\n").ToList();
        List<RvtElement> data = new();
        lines.ForEach(line =>
            {
                // ignore first line
                if (line.StartsWith(DateAndTimeString) || string.IsNullOrEmpty(line))
                {
                    return;
                }

                data.Add(new RvtElement(line));
            }
        );

        return data;
    }

    /// <summary>
    /// Create the output files.
    /// </summary>
    /// <param name="data">The input data from the exported file.</param>
    /// <param name="filePath">The path to store the finished report.</param>
    public static void CreateReport(List<RvtStatistics?> data, string filePath,bool csvIsEnabled, bool dailyReportIsEnabled)
    {
        string currentDateTime = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
        string fileName = "rvt-report-" + currentDateTime;
        string filePathWithoutExtension = Path.Combine(filePath, fileName);

        // create an empty directory if it does not exist
        Directory.CreateDirectory(filePath);

        if (CreateExcelOutput)
        {
            ExcelFunctions.CreateExcel(data, filePathWithoutExtension, dailyReportIsEnabled);
        }

        if (csvIsEnabled)
        {
            CsvFunctions.CreateCsvOutput(data, filePathWithoutExtension);
        }

        MessageBox.Show("Output file was created successfully.", "Report created", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    /// <summary>
    /// Read the content of a text file.
    /// If the input file is not in Unicode, it will be read as UTF-8.
    /// </summary>
    /// <param name="filePath">The path to the file to read.</param>
    /// <returns>A string containing the file content.</returns>
    private static string ReadTextFile(string filePath)
    {
        string fileContent = File.ReadAllText(filePath, Encoding.Unicode);
        return fileContent.Contains("\r\n") 
            ? fileContent 
            : File.ReadAllText(filePath, Encoding.UTF8);
    }
}