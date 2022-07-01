using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NXOpen;
using NXOpen.UF;

namespace TableTest
{
    class Sketch
    {
        public static void Main(string[] args){
        
            Session session = Session.GetSession();
            UFSession uFsession = UFSession.GetUFSession();
            Part part = session.Parts.Work;

            string name = "customSketch";
            uFsession.Sket.InitializeSketch(ref name, out Tag tag);

            double[] matrix = new double[9];
            DatumCollection datumCollection = part.Datums;
            Vector3d datumAxis3DX = new Vector3d();
            Vector3d datumAxis3DY = new Vector3d();

            foreach (var datum in datumCollection)
            {
                if(datum.GetType() == typeof(DatumAxis))
                {
                    DatumAxis datumAxis = (DatumAxis)datum;
                    if(datumAxis.Direction.X == 1.0)
                    {
                        datumAxis3DX = datumAxis.Direction;
                    }
                    if (datumAxis.Direction.Y == 1.0)
                    {
                        datumAxis3DY = datumAxis.Direction;
                    }
                }
            }

            if(datumAxis3DX.X == 1 && datumAxis3DX.Y == 1)
            {
                matrix[0] = datumAxis3DX.X;
                matrix[1] = datumAxis3DX.Y;
                matrix[2] = datumAxis3DX.Z;
                matrix[3] = datumAxis3DX.X;
                matrix[4] = datumAxis3DX.Y;
                matrix[5] = datumAxis3DX.Z;
                matrix[6] = 0.0;
                matrix[7] = 0.0;
                matrix[8] = 0.0;

                Tag[] tags = new Tag[0];
                int[] refs = new int[0];
                uFsession.Sket.CreateSketch(name, 2, matrix, tags, refs, 1, out Tag sketchTag);

                session.ActiveSketch.Activate(NXOpen.Sketch.ViewReorient.True);
            }
        }
    }
}
