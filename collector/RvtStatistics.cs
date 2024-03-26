using System;
using System.Collections.Generic;
using System.Linq;
using rvt_report_creator.measurements;

namespace rvt_report_creator.collector;

public class RvtStatistics
{
    private const int HoursPerDay = 24;
    private const double ReturnValueForEmptyList = double.NaN;

    public DateTime Date { get; set; }
    public List<RvtElement> Elements { get; set; } = new();

    /// <summary>
    /// Create a percentile value for the temperature.
    /// </summary>
    /// <param name="startTime">The start time of the data set.</param>
    /// <param name="endTime">The end time of the data set.</param>
    /// <param name="percentile">The percentile value from 0.0 to 1.0</param>
    /// <returns>The percentile of the measurements.</returns>
    public double CreateTemperaturePercentile(TimeSpan startTime, TimeSpan endTime, double percentile)
    {
        List<double> temperatures = (
            from element in Elements
            where element.DateAndTime.TimeOfDay >= startTime
                  && element.DateAndTime.TimeOfDay <= endTime
            select element.Mbw.Temperature
        ).ToList();

        if (temperatures.Count == 0)
        {
            return ReturnValueForEmptyList;
        }

        temperatures.Sort();
        int index = (int)(percentile * temperatures.Count);
        return temperatures[index];
    }

    /// <summary>
    /// Calculate the daily average flow rate of the container.
    /// </summary>
    /// <returns></returns>
    public double CalculateContainerOutflowPerDay()
    {
        if (Elements.Count == 0)
        {
            return ReturnValueForEmptyList;
        }

        return CalculateContainerOutflowMean() * HoursPerDay;
    }

    /// <summary>
    /// Calculate the mean outflow of the container.
    /// </summary>
    /// <returns></returns>
    public double CalculateContainerOutflowMean()
    {
        if (Elements.Count == 0)
        {
            return ReturnValueForEmptyList;
        }

        return Elements.Sum(element => element.Pufferbehaelter.Ablauf) / Elements.Count;
    }

    /// <summary>
    /// Calculate the daily average outflow to the water.
    /// </summary>
    /// <returns></returns>
    public double CalculateOutFlowPerDay()
    {
        if (Elements.Count == 0)
        {
            return ReturnValueForEmptyList;
        }

        return CalculateOutFlowMean() * HoursPerDay;
    }

    /// <summary>
    /// Calculate the mean outflow of the elements.
    /// </summary>
    /// <returns></returns>
    public double CalculateOutFlowMean()
    {
        if (Elements.Count == 0)
        {
            return ReturnValueForEmptyList;
        }

        return Elements.Sum(element => element.Mbw.Durchfluss) / Elements.Count;
    }

    /// <summary>
    /// Calculates the minimum pH value of the elements.
    /// </summary>
    /// <returns></returns>
    public double CalculatePhMin()
    {
        return Elements.Count == 0
            ? ReturnValueForEmptyList
            : Elements.Min(element => element.Mbw.PhWert);
    }

    /// <summary>
    /// Calculates the maximum pH value of the elements.
    /// </summary>
    /// <returns></returns>
    public double CalculatePhMax()
    {
        return Elements.Count == 0
            ? ReturnValueForEmptyList
            : Elements.Max(element => element.Mbw.PhWert);
    }

    /// <summary>
    /// Calculates the mean pH value of the elements.
    /// </summary>
    /// <returns></returns>
    public double CalculatePhMean()
    {
        if (Elements.Count == 0)
        {
            return ReturnValueForEmptyList;
        }

        double sum = Elements.Sum(element => element.Mbw.PhWert);
        return sum / Elements.Count;
    }

    /// <summary>
    /// Calculates the minimum temperature of the elements.
    /// </summary>
    /// <returns></returns>
    public double CalculateTemperatureMin()
    {
        return Elements.Count == 0
            ? ReturnValueForEmptyList
            : Elements.Min(element => element.Mbw.Temperature);
    }

    /// <summary>
    /// Calculates the maximum temperature of the elements.
    /// </summary>
    /// <returns></returns>
    public double CalculateTemperatureMax()
    {
        return Elements.Count == 0
            ? ReturnValueForEmptyList
            : Elements.Max(element => element.Mbw.Temperature);
    }

    /// <summary>
    /// Calculates the mean temperature of the elements.
    /// </summary>
    /// <returns></returns>
    public double CalculateTemperatureMean()
    {
        if (Elements.Count == 0)
        {
            return ReturnValueForEmptyList;
        }

        double sum = Elements.Sum(element => element.Mbw.Temperature);
        return sum / Elements.Count;
    }
}