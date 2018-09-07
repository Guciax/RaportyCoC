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
            double maxX = Math.Max(histogram.Select(x => x.Key).Max(), normalCurve.Select(x => x.Key).Max()) * 1.1;
            double maxY = histogram.Select(y => y.Value).Max()*1.1;
            double minX = Math.Min(histogram.Select(x => x.Key).Min(), normalCurve.Select(x => x.Key).Min())*0.9;
            double minY = histogram.Select(y => y.Value).Min()*0.9;
            double maxYNormal = normalCurve.Select(max => max.Value).Max() * 1.15;
            double mean = normalCurve.Aggregate((l, r) => l.Value > r.Value ? l : r).Key;


            chart.Series.Clear();
            chart.Legends.Clear();
            chart.ChartAreas.Clear();
            chart.Annotations.Clear();

            ChartArea ar = new ChartArea();
            ar.AxisX.MajorGrid.Enabled = false;
            ar.AxisY.MajorGrid.Enabled = false;
            ar.AxisY.Enabled = AxisEnabled.False;
            ar.AxisY2.Enabled = AxisEnabled.False;
            ar.AxisY2.MajorGrid.Enabled = false;
            ar.AxisY.LabelStyle.Enabled = false;
            ar.AxisY2.LabelStyle.Enabled = false;
            ar.AxisX.IsLabelAutoFit = false;
            ar.AxisX.Minimum = minX;
            ar.AxisX.Maximum = maxX;
            ar.AxisY.Minimum = 0;
            ar.AxisY.Maximum = maxY;
            ar.AxisY2.Maximum = maxYNormal;
            
            chart.ChartAreas.Add(ar);

            Series normalCurveSeries = new Series();
            normalCurveSeries.ChartType = SeriesChartType.Spline;
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

            double maxValue = histogram.Keys.Max();
            double minValue = histogram.Keys.Min();

            if (lslValue > 0)
            {
                if (minValue > lslValue)
                {
                    chart.ChartAreas[0].AxisX.Minimum = lslValue * 0.995;
                }
                else
                {
                    chart.ChartAreas[0].AxisX.Minimum = minValue * 0.995;
                }
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

                RectangleAnnotation lslLabel = new RectangleAnnotation();
                lslLabel.AnchorY = maxY;
                lslLabel.AnchorX = lslValue;
                lslLabel.AnchorDataPoint = histogramSeries.Points[0];
                lslLabel.Font = new Font("Arial Narrow", 9, FontStyle.Bold);
                lslLabel.AnchorOffsetY = -5;
                lslLabel.AnchorOffsetX = 0;
                lslLabel.AnchorAlignment = ContentAlignment.MiddleLeft;
                lslLabel.Text = "LSL";
                lslLabel.BackColor = Color.White;
                lslLabel.LineWidth = 0;
                chart.Annotations.Add(lslLabel);
            }
            else
            {
                chart.ChartAreas[0].AxisX.Minimum = minValue * 0.995;
            }

            if (uslValue > 0)
            {
                if (uslValue > maxValue)
                {
                    chart.ChartAreas[0].AxisX.Maximum =  uslValue * 1.005;
                }
                else
                {
                    chart.ChartAreas[0].AxisX.Maximum = maxValue * 1.005;
                }
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

                RectangleAnnotation uslLabel = new RectangleAnnotation();
                uslLabel.AnchorY = maxY;
                uslLabel.AnchorX = uslValue;
                uslLabel.AnchorDataPoint = histogramSeries.Points[histogramSeries.Points.Count - 1];
                uslLabel.Font = new Font("Arial Narrow", 9, FontStyle.Bold);
                uslLabel.AnchorOffsetY = -5;     // *
                uslLabel.AnchorAlignment = ContentAlignment.MiddleLeft;
                uslLabel.Text = "USL";
                uslLabel.BackColor = Color.White;
                uslLabel.LineWidth = 0;
                chart.Annotations.Add(uslLabel);
            }
            else
            {
                chart.ChartAreas[0].AxisX.Maximum = maxValue * 1.005;
            }

            if (ar.AxisX.Maximum-ar.AxisX.Minimum > 20)
            {
                ar.AxisX.LabelStyle.Format = "{0}";
            }
            else
            {
                ar.AxisX.LabelStyle.Format = "{0.0}";
            }

            VerticalLineAnnotation meanLine = new VerticalLineAnnotation();
            meanLine.X = mean;
            meanLine.LineColor = Color.Lime;
            meanLine.LineDashStyle = ChartDashStyle.Dash;
            meanLine.AxisX = chart.ChartAreas[0].AxisX;
            meanLine.ClipToChartArea = chart.ChartAreas[0].Name;
            meanLine.Name = "mean";
            meanLine.Width = 10;
            meanLine.AnchorX = lslValue;
            meanLine.IsInfinitive = true;
            chart.Annotations.Add(meanLine);

            chart.Series.Add(histogramSeries);
            chart.Series.Add(normalCurveSeries);
        }

        public static void DrawElipse(Chart chart, Dictionary<string, PcbTesterMeasurements> measurements, Dictionary<double, Tuple<double, double>> elipseBorder)
        {
            chart.Series.Clear();
            chart.ChartAreas.Clear();
            chart.Legends.Clear();

            ChartArea ar = new ChartArea();
            ar.AxisX.LabelStyle.Format = "{0.0000}";
            ar.AxisX.IsLabelAutoFit = false;
            ar.AxisY.LabelStyle.Format = "{0.0000}";
            ar.AxisY.IsLabelAutoFit = false;
            ar.AxisX.MajorGrid.Enabled = false;
            ar.AxisY.MajorGrid.Enabled = false;

            Series elipseBorderSeriesPuls = new Series();
            elipseBorderSeriesPuls.Color = Color.Red;
            elipseBorderSeriesPuls.ChartType = SeriesChartType.FastLine;
            elipseBorderSeriesPuls.BorderWidth = 3;

            Series elipseBorderSeriesMinus = new Series();
            elipseBorderSeriesMinus.Color = Color.Red;
            elipseBorderSeriesMinus.ChartType = SeriesChartType.FastLine;
            elipseBorderSeriesMinus.BorderWidth = 3;

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
            measurementsSeries.Color = Color.DarkSeaGreen;
            measurementsSeries.ChartType = SeriesChartType.Point;
            measurementsSeries.MarkerSize = 4;
            measurementsSeries.MarkerStyle = MarkerStyle.Circle;

            foreach (var pcb in measurements)
            {
                DataPoint pt = new DataPoint();
                pt.SetValueXY(pcb.Value.Cx, pcb.Value.Cy);
                measurementsSeries.Points.Add(pt);
            }

            double max = ar.AxisY.Maximum;
            double min = ar.AxisY.Minimum;

            ar.AxisY.Maximum = elipseBorder.Select(val => val.Value.Item1).Max()*1.01;
            ar.AxisY.Minimum = elipseBorder.Select(val => val.Value.Item2).Min()*0.99;

            chart.ChartAreas.Add(ar);
            chart.Series.Add(measurementsSeries);
            chart.Series.Add(elipseBorderSeriesPuls);
            chart.Series.Add(elipseBorderSeriesMinus);
        }

        public static Bitmap ConvertChartToBmp(Chart chart)
        {
            Bitmap result;
            
            using (MemoryStream ms = new MemoryStream())
            {
                chart.SaveImage(ms, ChartImageFormat.Bmp);
                result = new Bitmap(ms);
            }

            chart.DrawToBitmap(result, new Rectangle(0, 0, result.Width, result.Height));

            //chart.SaveImage(@"H:\" + chart.Name+".png", ChartImageFormat.Png);
            return result;
        }
    }
}
