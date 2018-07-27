using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RaportyCoC
{
    class Histogram
    {
        public static SortedDictionary<double, int> CalculateHistogram(List<double> values)
        {
            
            SortedDictionary<double, int> result = new SortedDictionary<double, int>();
            int stepsCount = 21;
            var min = values.Min();
            var max = values.Max();
            double stepWidth = ( max - min) / (double)stepsCount;
            List<double> steps = new List<double>();

            for (int i = 0; i < stepsCount; i++) 
            {
                double step = values.Min() + i * stepWidth;
                steps.Add(step);
                result.Add(step, 0);
            }

            foreach (var value in values)
            {
                double step = FindClosestStep(value, steps.ToArray());
                result[step]++;
            }

            return result;
        }

        public static SortedDictionary<double, double> CalculateHistogramNormal(SortedDictionary<double, int> histogram, List<double> values)
        {
            SortedDictionary<double, double> result = new SortedDictionary<double, double>();

            double mean = values.Average();
            double sumOfSquaresOfDifferences = values.Select(val => (val - mean) * (val - mean)).Sum();
            double stdDev = Math.Sqrt(sumOfSquaresOfDifferences / values.Count);
            double deviationSq = values.Select(x => Math.Pow(x - mean, 2)).Average();

            foreach (var x in histogram)
            {
                result.Add(x.Key, F(x.Key, mean, stdDev));
            }

            return result;
        }

        private static double FindClosestStep(double value, double[]steps)
        {

            var closest = steps.Aggregate((x, y) => Math.Abs(x - value) < Math.Abs(y - value) ? x : y);

            return closest;
        }

        // The normal distribution function.
        private static double F(double x, double mean, double stddev)
        {
            double var = Math.Pow(stddev, 2);
            return ((0.5 / Math.PI) * Math.Exp(-(x - mean) * (x - mean) / (2 * var)));
        }
    }
}
