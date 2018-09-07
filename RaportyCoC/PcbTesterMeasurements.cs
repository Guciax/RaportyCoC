using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RaportyCoC
{
    class PcbTesterMeasurements
    {
        public PcbTesterMeasurements(double Cx, double Cy, double Sdcm, double Cct, DateTime inspectionTime, string model, double vf, double lm, double lmW, double cri, double cct, string orderNo)
        {
            this.Cx = Cx;
            this.Cy = Cy;
            this.Sdcm = Sdcm;
            this.Cct = Cct;
            InspectionTime = inspectionTime;
            Model = model;
            Vf = vf;
            Lm = lm;
            LmW = lmW;
            Cri = cri;
            Cct1 = cct;
            OrderNo = orderNo;
        }

        public double Cx { get; }
        public double Cy { get; }
        public double Sdcm { get; }
        public double Cct { get; }
        public DateTime InspectionTime { get; }
        public string Model { get; }
        public double Vf { get; }
        public double Lm { get; }
        public double LmW { get; }
        public double Cri { get; }
        public double Cct1 { get; }
        public string OrderNo { get; }
    }
}
