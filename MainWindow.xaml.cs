using System.Collections.Generic;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using rvt_report_creator.collector;
using rvt_report_creator.files;
using rvt_report_creator.measurements;

namespace rvt_report_creator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Search for a file to use as input for the report.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SearchInputFileClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new();
            openFileDialog.Filter = "Text Dateien (*.txt;*.csv)|*.txt;*.csv|Alle Dateien (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                InputFilePath.Text = openFileDialog.FileName;
            }
        }

        /// <summary>
        /// Search for a file to use as input for the report.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SearchOutputFilePathClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == true)
            {
                OutputFilePath.Text = openFileDialog.FileName;
            }
        }

        /// <summary>
        /// Close the application.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CloseButtonClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /// <summary>
        /// Load the file from the given file path and create a report from it.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void StartButtonClick(object sender, RoutedEventArgs e)
        {
            try
            {
                CheckGivenFilePath();
            }
            catch (IOException exception)
            {
                MessageBox.Show(exception.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // check the state of the checkboxes for the output files
            bool csvIsEnabled = EnableCsvExport.IsChecked == true;
            bool dailyReportIsEnabled = EnableDailyReports.IsChecked == true;
            
            List<RvtElement> rvtElements = FileHandler.ReadFile(InputFilePath.Text);
            List<RvtStatistics?> rvtStatistics = FilterElementsByDate(rvtElements);
            FileHandler.CreateReport(
                rvtStatistics,
                OutputFilePath.Text,
                csvIsEnabled, 
                dailyReportIsEnabled
            );
        }

        /// <summary>
        /// Filter the elements by its date and store them into a list of statistic elements.
        /// </summary>
        /// <param name="rvtElements">Data from the export file.</param>
        /// <returns>A list of statistic elements, to load the values for the final report.</returns>
        private static List<RvtStatistics?> FilterElementsByDate(List<RvtElement> rvtElements)
        {
            List<RvtStatistics?> rvtStatistics = new();

            foreach (RvtElement rvtElement in rvtElements)
            {
                RvtStatistics? rvtStatistic = rvtStatistics.Find(statistic =>
                    statistic != null && statistic.Date.Date == rvtElement.DateAndTime.Date
                );

                if (rvtStatistic == null)
                {
                    rvtStatistic = new RvtStatistics
                    {
                        Date = rvtElement.DateAndTime.Date,
                        Elements = new List<RvtElement>()
                    };
                    rvtStatistics.Add(rvtStatistic);
                }

                rvtStatistic.Elements.Add(rvtElement);
            }

            return rvtStatistics;
        }

        /// <summary>
        /// Check the given file path for validity.
        /// </summary>
        /// <exception cref="FileNotFoundException">No file selected or selected file does not exist.</exception>
        private void CheckGivenFilePath()
        {
            if (InputFilePath == null || InputFilePath.Text == "")
            {
                throw new IOException(
                    "Es wurde keine Datei ausgewählt, die zur Erstellung des Reports verwendet wird. Bitte geben Sie zuerst einen Dateipfad der Export-Datei an.");
            }

            if (!File.Exists(InputFilePath.Text))
            {
                throw new FileNotFoundException(
                    "Die angegebene Datei konnte nicht gefunden werden. Bitte gehen Sie sicher, dass die Datei existiert!"
                );
            }
        }
    }
}