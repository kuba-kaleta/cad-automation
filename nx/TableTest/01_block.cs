using NXOpen;

namespace TableTest
{
    public class Block //NX open C# Create Block - dziala
    {
        public static void Main(string[] args)
        {
            Session session = Session.GetSession();

            Part workPart = session.Parts.Work;

            NXOpen.Features.BlockFeatureBuilder blockFeatureBuilder = null;
            NXOpen.Features.Feature feature = null;

            blockFeatureBuilder = workPart.Features.CreateBlockFeatureBuilder(feature);

            blockFeatureBuilder.BooleanOption.Type = NXOpen.GeometricUtilities.BooleanOperation.BooleanType.Create;
            blockFeatureBuilder.Type = NXOpen.Features.BlockFeatureBuilder.Types.OriginAndEdgeLengths;

            Point3d origin = new Point3d(0.0, 0.0, 0.0);
            Point originPoint = workPart.Points.CreatePoint(origin);

            blockFeatureBuilder.OriginPoint = originPoint;
            blockFeatureBuilder.SetOriginAndLengths(origin, "100", "200", "500");

            blockFeatureBuilder.CommitFeature();
            blockFeatureBuilder.Destroy();
        }

        public static int GetUnloadOption(string arg)
        {
            return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);
        }
    }
}

