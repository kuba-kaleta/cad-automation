using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

using NXOpen;
using NXOpen.UF;

namespace TableTest
{
    class Revolve //Create Revolve Feature From NX Open - dziala
    {
        public static void Main(string[] args)
        {
            Session session = Session.GetSession();
            UFSession uFSession = UFSession.GetUFSession();
            UFCurve.Line line = new UFCurve.Line();

            line.start_point = new double[3];
            line.start_point[0] = 0.0;
            line.start_point[1] = 0.0;
            line.start_point[2] = 0.0;

            line.end_point = new double[3];
            line.end_point[0] = 0.0;
            line.end_point[1] = 0.0;
            line.end_point[2] = 100.0;
            
            Tag newLine;
            uFSession.Curve.CreateLine(ref line, out newLine);

            double[] direction = new double[] { 0.0, 0.0, 1.0 };
            Tag[] objects = new Tag[] { newLine };

            string[] limits = new string[2];
            limits[0] = "0.0";
            limits[1] = "360.0";
            double[] origin = new double[] { 10, 0.0, 0.0 };

            Tag[] feats;
            uFSession.Modl.CreateRevolved(objects, limits, origin, direction, FeatureSigns.Nullsign, out feats);

        }

        public static int GetUnloadOption(string arg)
        {
            return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);
        }

    }
}
