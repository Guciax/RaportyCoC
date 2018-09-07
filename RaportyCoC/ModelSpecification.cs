using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RaportyCoC
{
    class ModelSpecification
    {
        public ModelSpecification(double maxSdcm, double Cx, double Cy, double A, double B, double C, double D, double theta, double variable, double cctMin, double cctMax, double Vf_min, double Vf_max, double lm_min, double lm_max, double lmW_min, double CRI_min, double CRI_max, double currentForward, string tridonicCustomerNumner, string tridonicDescription, string lgitName, string lgitDescription, string[] boxes, string orderNo)

        {
            MaxSdcm = maxSdcm;
            this.Cx = Cx;
            this.Cy = Cy;
            this.A = A;
            this.B = B;
            this.C = C;
            this.D = D;
            Theta = theta;
            Variable = variable;
            CctMin = cctMin;
            CctMax = cctMax;
            Vf_Min = Vf_min;
            Vf_Max = Vf_max;
            Lm_Min = lm_min;
            Lm_Max = lm_max;
            LmW_Min = lmW_min;
            CRI_Min = CRI_min;
            CRI_Max = CRI_max;
            CurrentForward = currentForward;
            TridonicCustomerNumner = tridonicCustomerNumner;
            TridonicDescription = tridonicDescription;
            LgitName = lgitName;
            LgitDescription = lgitDescription;
            Boxes = boxes;
            OrderNo = orderNo;
        }

        public double MaxSdcm { get; }
        public double Cx { get; }
        public double Cy { get; }
        public double A { get; }
        public double B { get; }
        public double C { get; }
        public double D { get; }
        public double Theta { get; }
        public double Variable { get; }
        public double CctMin { get; set; }
        public double CctMax { get; set; }
        public double Vf_Min { get; set; }
        public double Vf_Max { get; set; }
        public double Lm_Min { get; set; }
        public double Lm_Max { get; set; }
        public double LmW_Min { get; set; }
        public double CRI_Min { get; set; }
        public double CRI_Max { get; set; }
        public double CurrentForward { get; set; }
        public string TridonicCustomerNumner { get; }
        public string TridonicDescription { get; }
        public string LgitName { get; }
        public string LgitDescription { get; }
        public string[] Boxes { get; }
        public string OrderNo { get; set; }
    }
}
