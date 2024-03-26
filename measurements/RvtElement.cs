using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace rvt_report_creator.measurements;

public class RvtElement
{
      private readonly CultureInfo _cultureInfo = CultureInfo.InvariantCulture;
    private const string Format = "HH:mm:ss.fff dd-MM-yyyy";

    public DateTime DateAndTime { get; set; }
    public string Message { get; set; }
    public Sensors Eingaenge { get; set; }
    public Container Pufferbehaelter { get; set; }
    public Measurements Mbw { get; set; }
    public Sewage Abwasser { get; set; }

    public RvtElement(string line)
    {
        List<string?> elements = line.Split(";").ToList();

        if (elements.Count == 1)
        {
            return;
        }

        string? dateString = elements[0];

        DateAndTime = DateTime.ParseExact(dateString, Format, _cultureInfo);

        Message = elements[1];
        Eingaenge = new Sensors
        {
            L301 = ParseDoubleValue(elements[2]),
            T101 = ParseDoubleValue(elements[8]),
            F206 = ParseDoubleValue(elements[9]),
            F103 = ParseDoubleValue(elements[10]),
            L203 = ParseDoubleValue(elements[11]),
            Q220 = ParseDoubleValue(elements[12]),
            Q221 = ParseDoubleValue(elements[13]),
        };

        Pufferbehaelter = new Container
        {
            Durchfluss = ParseDoubleValue(elements[3]),
            Ablauf = ParseDoubleValue(elements[14])
        };

        Mbw = new Measurements
        {
            Durchfluss = ParseDoubleValue(elements[4]),
            Temperature = ParseDoubleValue(elements[5]),
            Leitfaehigkeit = ParseDoubleValue(elements[6]),
            PhWert = ParseDoubleValue(elements[7])
        };

        Abwasser = new Sewage
        {
            Leitmenge = ParseDoubleValue(elements[15]),
            Temperatur = new Range
            {
                Minimum = ParseDoubleValue(elements[17]),
                Maximum = ParseDoubleValue(elements[16])
            },
            PhWert = new Range
            {
                Minimum = ParseDoubleValue(elements[19]),
                Maximum = ParseDoubleValue(elements[18])
            }
        };
    }

    private double ParseDoubleValue(string? element)
    {
        return double.Parse(element.Replace(',', '.'), _cultureInfo);
    }
}