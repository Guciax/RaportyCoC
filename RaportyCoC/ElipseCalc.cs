using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RaportyCoC
{
    class ElipseCalc
    {
        Dictionary<double, Tuple<double, double>> CalculateElipsePoints(Dictionary<string, PcbTesterMeasurements> pcbMeasurements, ModelSpecification spec)
        {
            Dictionary<double, Tuple<double, double>> result = new Dictionary<double, Tuple<double, double>>();
            double thetaRad = (Math.PI / 180) * spec.Theta;
            double a = 1 / Math.Pow(spec.C, 2) * Math.Pow(Math.Sin(thetaRad), 2) + 1 / Math.Pow(spec.D, 2) * Math.Pow(Math.Cos(thetaRad), 2);
            double g11 = 1 / Math.Pow(spec.C, 2) * Math.Pow(Math.Cos(thetaRad), 2) + 1 / Math.Pow(spec.D, 2) * Math.Pow(Math.Sin(thetaRad), 2);
            double xRange = spec.Variable * spec.C * spec.MaxSdcm * Math.Cos(thetaRad);

            for (int i = -50; i < 51; i++)
            {
                double dx = ((i == 50 || i == -50) ? i * xRange / 102 : i * xRange / 100);
                double X = spec.Cx + dx;

                double b = 2 * dx * (1 / Math.Pow(spec.C, 2) - 1 / Math.Pow(spec.D, 2)) * Math.Sin(thetaRad) * Math.Cos(thetaRad);
                double c = g11 * Math.Pow(dx, 2) - Math.Pow(spec.MaxSdcm, 2);
                double dYplus = -b + Math.Sqrt(Math.Pow(b, 2) - 4 * a * c) / (2 * a);
                double dYminus = -b - Math.Sqrt(Math.Pow(b, 2) - 4 * a * c) / (2 * a);
                double Yplus = spec.Cy + dYplus;
                double Yminus = spec.Cy + dYminus;

                result.Add(X, new Tuple<double, double>(Yplus, Yminus));
            }

            return result;
        }
    }
}
