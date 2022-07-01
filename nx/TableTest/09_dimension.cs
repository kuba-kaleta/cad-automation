using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NXOpen;
using NXOpen.UF;
using NXOpen.Drawings;

namespace TableTest
{
    public class Dimension //How to Create a Dimension using NX open - niedokonczone
    {
        static NXOpen.Session theSession = NXOpen.Session.GetSession();
        static UFSession ufSession = UFSession.GetUFSession();
        static NXOpen.Part workPart = theSession.Parts.Work;

        public static void Main()
        {
            //Check the module you run your application 
            ufSession.UF.AskApplicationModule(out int appModule);
            theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING");
            workPart.Drafting.EnterDraftingApplication();

            NXOpen.Drawings.DraftingDrawingSheet nullNXOpen_Drawings_DraftingDrawingSheet = null;
            NXOpen.Drawings.DraftingDrawingSheetBuilder draftingDrawingSheetBuilder1;

            draftingDrawingSheetBuilder1 = workPart.DraftingDrawingSheets.CreateDraftingDrawingSheetBuilder(nullNXOpen_Drawings_DraftingDrawingSheet);
            draftingDrawingSheetBuilder1.AutoStartViewCreation = true;
            draftingDrawingSheetBuilder1.Option = NXOpen.Drawings.DrawingSheetBuilder.SheetOption.StandardSize;
            draftingDrawingSheetBuilder1.StandardMetricScale = NXOpen.Drawings.DrawingSheetBuilder.SheetStandardMetricScale.S11;
            draftingDrawingSheetBuilder1.StandardEnglishScale = NXOpen.Drawings.DrawingSheetBuilder.SheetStandardEnglishScale.S11;
            draftingDrawingSheetBuilder1.AutoStartViewCreation = false;
            draftingDrawingSheetBuilder1.MetricSheetTemplateLocation = "C:\\Program Files\\Siemens\\NX 12.0\\DRAFTING\\templates\\Drawing-A0-Size2D-template.prt";
            draftingDrawingSheetBuilder1.Height = 841.0;
            draftingDrawingSheetBuilder1.Length = 1189.0;
            draftingDrawingSheetBuilder1.StandardMetricScale = NXOpen.Drawings.DrawingSheetBuilder.SheetStandardMetricScale.S11;
            draftingDrawingSheetBuilder1.ScaleNumerator = 1.0;
            draftingDrawingSheetBuilder1.ScaleDenominator = 1.0;
            draftingDrawingSheetBuilder1.Units = NXOpen.Drawings.DrawingSheetBuilder.SheetUnits.Metric;
            draftingDrawingSheetBuilder1.ProjectionAngle = NXOpen.Drawings.DrawingSheetBuilder.SheetProjectionAngle.Third;
            draftingDrawingSheetBuilder1.Number = "1";
            draftingDrawingSheetBuilder1.SecondaryNumber = "";
            draftingDrawingSheetBuilder1.Revision = "A";
            draftingDrawingSheetBuilder1.MetricSheetTemplateLocation = "C:\\Program Files\\Siemens\\NX 12.0\\DRAFTING\\templates\\Drawing-A0-Size2D-template.prt";

            NXOpen.NXObject nXObject1;
            nXObject1 = draftingDrawingSheetBuilder1.Commit();

            draftingDrawingSheetBuilder1.Destroy();
            ufSession.Draw.AskCurrentDrawing(out Tag drawTag);

            //Important Check Here
            if (drawTag != Tag.Null)
            {
                CreateDrawingView(drawTag);
            }
        }

        public static void CreateDrawingView(Tag currrentDrawingTag)
        {
            try
            {
                //Create View For Top view From current NX workpart
                ModelingViewCollection mviews = workPart.ModelingViews;
                NXOpen.ModelingView topView = null;
                foreach (ModelingView mv in mviews)
                {
                    if (mv.Name.Equals("Top"))
                    {
                        topView = mv;
                        break;
                    }
                }
                DrawingSheet dSheet = workPart.DrawingSheets.CurrentDrawingSheet;
                //Set View Placement point
                NXOpen.Point3d point1 = new NXOpen.Point3d(dSheet.Length / 2, dSheet.Height / 2, 0.0);

                double[] viewCenter = new double[2];
                viewCenter[0] = point1.X;
                viewCenter[1] = point1.Y;

                UFDraw.ViewInfo viewInfo = new UFDraw.ViewInfo();
                viewInfo.anchor_point = Tag.Null;
                viewInfo.arrangement_name = "A";
                viewInfo.model_name = "Test";
                viewInfo.use_ref_pt = false;
                viewInfo.view_scale = 1.0;
                viewInfo.transfer_annotation = true;
                viewInfo.inherit_boundary = false;
                viewInfo.view_status = UFDraw.ViewStatus.ActiveView;

                ufSession.Draw.ImportView(currrentDrawingTag, topView.Tag, viewCenter, ref viewInfo, out Tag draw_view_tag);
                //ufSession.Draw.SetAutoUpdate(draw_view_tag, out bool autoUpdate);
                DraftingView[] allViews = workPart.DrawingSheets.CurrentDrawingSheet.GetDraftingViews();
                workPart.DraftingViews.UpdateViews(allViews);

                CreateDiametricDimension(dSheet, allViews[0]);
            }
            catch (Exception exception)
            {
                Console.WriteLine("Error");
            }
        }

        public static void CreateDiametricDimension(DrawingSheet drawingSheet, DraftingView draftingView)
        {
            DisplayableObject[] displayableObjects = draftingView.AskVisibleObjects();

            foreach(var obj in displayableObjects)
            {
                //Console.WriteLine(obj.GetType());
                if(obj.GetType() == typeof(DraftingCurve))
                {
                    ufSession.Obj.AskExtendedTypeAndSubtype(obj.Tag, out int type, out int subtype);

                    if (type == UFConstants.UF_circle_type)
                    {
                        UFDrf.Object ufDrfObj = new UFDrf.Object();
                        UFDrf.Text text = new UFDrf.Text();
                        ufDrfObj.object_tag = obj.Tag;
                        ufDrfObj.object_view_tag = draftingView.Tag;
                        ufDrfObj.object_assoc_type = UFDrf.AssocType.Intersection;
                        double[] dimOrigin = new double[3];
                        dimOrigin[0] = drawingSheet.Length / 2 + 100.0;
                        dimOrigin[1] = drawingSheet.Height / 2 + 100.0;
                        dimOrigin[2] = 0.0;
                        ufSession.Drf.CreateDiameterDim(ref ufDrfObj, ref text, dimOrigin, out Tag dimension);
                    }
                }
            }
        }

        public static int GetUnloadOption(string dummy)
        {
            return (int)NXOpen.Session.LibraryUnloadOption.Immediately;
        }

    }
}
