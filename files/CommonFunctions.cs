using System;
using System.Collections.Generic;

namespace rvt_report_creator.files;

public abstract class CommonFunctions
{
     public const double DefaultPercentage = 0.80;
    
    /// <summary>
    /// Load the header elements for the Excel file.
    /// </summary>
    /// <returns>A list of the heading elements.</returns>
    public static List<string> LoadHeaderElements()
    {
        List<string> headerElements = new()
        {
            "Monat",
            "Tag",
            "Auslaufmenge Pufferbehälter",
            "Einleitmenge",
            "Tagesmaximum pH-Wert",
            "Tagesminimum pH-Wert",
            "Tagesmittelwert Temperatur",
            "Tagesmaxumalwert Temperatur",
            "Temperatur Perzentil 0.8 - 00:00h-05:59h",
            "Temperatur Perzentil 0.8 - 06:00h-11:59h",
            "Temperatur Perzentil 0.8 - 12:00h-17:59h",
            "Temperatur Perzentil 0.8 - 18:00h-23:59h"
        };
        return headerElements;
    }

    /// <summary>
    /// Load the sub header units for the Excel file.
    /// </summary>
    /// <returns>A list of all units.</returns>
    public static List<string> LoadSubHeaderElements()
    {
        List<string> subHeaderElements = new()
        {
            "",
            "",
            "m³/d",
            "m³/d",
            "",
            "",
            "°C",
            "°C",
            "°C",
            "°C",
            "°C",
            "°C"
        };
        return subHeaderElements;
    }
    
    /// <summary>
    /// Create the start times for the required percentiles.
    /// </summary>
    /// <returns>A list of start times.</returns>
    public static List<TimeSpan> LoadStartTime()
    {
        List<TimeSpan> startTime = new()
        {
            new TimeSpan(0, 0, 0),
            new TimeSpan(6, 0, 0),
            new TimeSpan(12, 0, 0),
            new TimeSpan(18, 0, 0)
        };
        return startTime;
    }

    /// <summary>
    /// Create the end times for the required percentiles.
    /// </summary>
    /// <returns>A list of end times.</returns>
    public static List<TimeSpan> LoadEntTime()
    {
        List<TimeSpan> endTimes = new()
        {
            new TimeSpan(5, 59, 59),
            new TimeSpan(11, 59, 59),
            new TimeSpan(17, 59, 59),
            new TimeSpan(23, 59, 59)
        };
        return endTimes;
    }
}