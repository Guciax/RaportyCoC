using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.DataVisualization.Charting;

namespace RaportyCoC
{
    class Charting
    {
        public static void DrawHistogramChart(Chart chart, List<double> values, double lslValue, double uslValue)
        {
            SortedDictionary<double, int> histogram = Histogram.CalculateHistogram(values);
            SortedDictionary<double, double> normalCurve = Histogram.CalculateHistogramNormal(histogram, values);

            chart.Series.Clear();
            chart.Legends.Clear();
            chart.ChartAreas.Clear();

            ChartArea ar = new ChartArea();
            ar.AxisX.MajorGrid.Enabled = false;
            ar.AxisY.MajorGrid.Enabled = false;
            ar.AxisY.Enabled = AxisEnabled.False;
            ar.AxisY2.Enabled = AxisEnabled.False;
            ar.AxisY2.MajorGrid.Enabled = false;
            ar.AxisY.LabelStyle.Enabled = false;
            ar.AxisY2.LabelStyle.Enabled = false;
            ar.AxisX.LabelStyle.Format = "{0}";
            chart.ChartAreas.Add(ar);

            Series normalCurveSeries = new Series();
            normalCurveSeries.ChartType = SeriesChartType.FastLine;
            normalCurveSeries.BorderDashStyle = ChartDashStyle.Dash;
            normalCurveSeries.BorderWidth = 1;
            normalCurveSeries.Color = Color.Black;
            normalCurveSeries.YAxisType = AxisType.Secondary;

            Series histogramSeries = new Series();
            histogramSeries.ChartType = SeriesChartType.Column;
            histogramSeries.BorderColor = Color.DarkGray;
            histogramSeries.Color = Color.Silver;

            foreach (var xEntry in histogram)
            {
                DataPoint pt = new DataPoint();
                pt.SetValueXY(xEntry.Key, xEntry.Value);
                histogramSeries.Points.Add(pt);
            }

            foreach (var xEntry in normalCurve)
            {
                DataPoint pt = new DataPoint();
                pt.SetValueXY(xEntry.Key, xEntry.Value);
                normalCurveSeries.Points.Add(pt);
            }


                chart.ChartAreas[0].AxisX.Minimum = lslValue - lslValue * 0.005;


                chart.ChartAreas[0].AxisX.Maximum = uslValue + uslValue * 0.005;

            if (lslValue > 0)
            {
                VerticalLineAnnotation lsl = new VerticalLineAnnotation();
                lsl.X = lslValue;
                lsl.LineColor = Color.Red;
                lsl.LineDashStyle = ChartDashStyle.Dash;
                lsl.AxisX = chart.ChartAreas[0].AxisX;
                lsl.ClipToChartArea = chart.ChartAreas[0].Name;
                lsl.Name = "LSL";
                lsl.Width = 5;
                lsl.AnchorX = lslValue;
                lsl.IsInfinitive = true;
                chart.Annotations.Add(lsl);
            }

            if (uslValue > 0)
            {
                VerticalLineAnnotation usl = new VerticalLineAnnotation();
                usl.X = uslValue;
                usl.LineColor = Color.Red;
                usl.LineDashStyle = ChartDashStyle.Dash;
                usl.AxisX = chart.ChartAreas[0].AxisX;
                usl.ClipToChartArea = chart.ChartAreas[0].Name;
                usl.Name = "USL";
                usl.Width = 5;
                usl.AnchorX = uslValue;
                usl.IsInfinitive = true;
                usl.SmartLabelStyle.Enabled = true;
                chart.Annotations.Add(usl);
            }

            chart.Series.Add(histogramSeries);
            chart.Series.Add(normalCurveSeries);
        }

        public static void DrawElipse(Chart chart, Dictionary<string, PcbTesterMeasurements> measurements, Dictionary<double, Tuple<double, double>> elipseBorder)
        {
            chart.Series.Clear();
            chart.ChartAreas.Clear();
            chart.Legends.Clear();

            ChartArea ar = new ChartArea();

            Series elipseBorderSeriesPuls = new Series();
            elipseBorderSeriesPuls.Color = Color.Red;

            Series elipseBorderSeriesMinus = new Series();

            foreach (var point in elipseBorder)
            {
                DataPoint ptPlus = new DataPoint();
                ptPlus.SetValueXY(point.Key, point.Value.Item1);
                elipseBorderSeriesPuls.Points.Add(ptPlus);

                DataPoint ptMinus = new DataPoint();
                ptMinus.SetValueXY(point.Key, point.Value.Item2);
                elipseBorderSeriesMinus.Points.Add(ptMinus);
            }

            Series measurementsSeries = new Series();
            measurementsSeries.ChartType = SeriesChartType.Point;
            measurementsSeries.MarkerSize = 3;
            measurementsSeries.MarkerStyle = MarkerStyle.Cross;

            foreach (var pcb in measurements)
            {
                DataPoint pt = new DataPoint();
                pt.SetValueXY(pcb.Value.Cx, pcb.Value.Cy);
                measurementsSeries.Points.Add(pt);
            }

            chart.Series.Add(measurementsSeries);
            chart.Series.Add(elipseBorderSeriesPuls);
            chart.Series.Add(elipseBorderSeriesMinus);
        }

        public static Bitmap ConvertChartToBmp(Chart chart)
        {
            Bitmap result;
            using (MemoryStream ms = new MemoryStream())
            {
                chart.SaveImage(ms, ChartImageFormat.Jpeg);
                result = new Bitmap(ms);
            }
            return result;
        }
    }
}
