// NX 1988
// Journal created by jakub.krajanowski on Tue Oct 26 07:51:12 2021 Środkowoeuropejski czas letni

//
using System;
using System.Collections;
using NXOpen;

using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

public class NXJournal
{
    public static void Main(string[] args)
    {
        UI theUI = UI.GetUI();

        string path = @"D:\jakub.kaleta\NX\wymiary.xlsx";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;
        wb = excel.Workbooks.Open(path);
        ws = wb.Worksheets[1];
        double war;
        string warS;

        NXOpen.Session theSession = NXOpen.Session.GetSession();
        NXOpen.Part workPart = theSession.Parts.Work;
        NXOpen.Part displayPart = theSession.Parts.Display;
        // ----------------------------------------------
        //   Menu: Edit->Move Object...
        // ----------------------------------------------

        //warS = (string)ws.Cells[2, 1].Value;
        //theUI.NXMessageBox.Show("Informancje", NXMessageBox.DialogType.Information, warS);

        for (int i = 2; i <= 78; i++)
        {
            if (ws.Cells[i, 1] != null && ws.Cells[i, 1].Value2 != null && ws.Cells[i, 4].Value2.ToString() == "true")
            {
                warS = ws.Cells[i, 1].Value2.ToString();
                //theUI.NXMessageBox.Show("Informancje", NXMessageBox.DialogType.Information, warS);
                war = Convert.ToDouble(warS);

                NXOpen.Session.UndoMarkId markId1;
                markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Start");

                NXOpen.Features.MoveObject nullNXOpen_Features_MoveObject = null;
                NXOpen.Features.MoveObjectBuilder moveObjectBuilder1;
                moveObjectBuilder1 = workPart.BaseFeatures.CreateMoveObjectBuilder(nullNXOpen_Features_MoveObject);
                moveObjectBuilder1.TransformMotion.DistanceAngle.OrientXpress.AxisOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Axis.Passive;
                moveObjectBuilder1.TransformMotion.DistanceAngle.OrientXpress.PlaneOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Plane.Passive;
                moveObjectBuilder1.TransformMotion.AlongCurveAngle.AlongCurve.IsPercentUsed = true;
                moveObjectBuilder1.TransformMotion.AlongCurveAngle.AlongCurve.Expression.SetFormula("0");
                moveObjectBuilder1.TransformMotion.AlongCurveAngle.AlongCurve.Expression.SetFormula("0");
                moveObjectBuilder1.TransformMotion.OrientXpress.AxisOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Axis.Passive;
                moveObjectBuilder1.TransformMotion.OrientXpress.PlaneOption = NXOpen.GeometricUtilities.OrientXpressBuilder.Plane.Passive;
                moveObjectBuilder1.TransformMotion.DeltaEnum = NXOpen.GeometricUtilities.ModlMotion.Delta.ReferenceAcsWorkPart;
                moveObjectBuilder1.TransformMotion.Option = NXOpen.GeometricUtilities.ModlMotion.Options.Dynamic;
                moveObjectBuilder1.TransformMotion.DistanceValue.SetFormula("0");
                moveObjectBuilder1.TransformMotion.DistanceBetweenPointsDistance.SetFormula("0");
                moveObjectBuilder1.TransformMotion.RadialDistance.SetFormula("0");
                moveObjectBuilder1.TransformMotion.Angle.SetFormula("0");
                moveObjectBuilder1.TransformMotion.DistanceAngle.Distance.SetFormula("0");
                moveObjectBuilder1.TransformMotion.DistanceAngle.Angle.SetFormula("0");
                moveObjectBuilder1.TransformMotion.DeltaXc.SetFormula("0");
                moveObjectBuilder1.TransformMotion.DeltaYc.SetFormula("0");
                moveObjectBuilder1.TransformMotion.DeltaZc.SetFormula("0");
                moveObjectBuilder1.TransformMotion.AlongCurveAngle.AlongCurve.Expression.SetFormula("0");
                moveObjectBuilder1.TransformMotion.AlongCurveAngle.AlongCurveAngle.SetFormula("0");
                moveObjectBuilder1.MoveObjectResult = NXOpen.Features.MoveObjectBuilder.MoveObjectResultOptions.CopyOriginal;
                theSession.SetUndoMarkName(markId1, "Move Object Dialog");

                NXOpen.NXObject[] objects1 = new NXOpen.NXObject[9];
                NXOpen.Body body1 = ((NXOpen.Body)workPart.Bodies.FindObject("COMBINE_SHEETS(80)"));
                objects1[0] = body1;
                NXOpen.Body body2 = ((NXOpen.Body)workPart.Bodies.FindObject("COMBINE_SHEETS(80)1"));
                objects1[1] = body2;
                NXOpen.Body body3 = ((NXOpen.Body)workPart.Bodies.FindObject("COMBINE_SHEETS(80)2"));
                objects1[2] = body3;
                NXOpen.Body body4 = ((NXOpen.Body)workPart.Bodies.FindObject("COMBINE_SHEETS(80)3"));
                objects1[3] = body4;
                NXOpen.Body body5 = ((NXOpen.Body)workPart.Bodies.FindObject("COMBINE_SHEETS(80)4"));
                objects1[4] = body5;
                NXOpen.Body body6 = ((NXOpen.Body)workPart.Bodies.FindObject("COMBINE_SHEETS(80)5"));
                objects1[5] = body6;
                NXOpen.Body body7 = ((NXOpen.Body)workPart.Bodies.FindObject("COMBINE_SHEETS(80)6"));
                objects1[6] = body7;
                NXOpen.Body body8 = ((NXOpen.Body)workPart.Bodies.FindObject("COMBINE_SHEETS(80)7"));
                objects1[7] = body8;
                NXOpen.Body body9 = ((NXOpen.Body)workPart.Bodies.FindObject("COMBINE_SHEETS(80)8"));
                objects1[8] = body9;
                bool added1;
                added1 = moveObjectBuilder1.ObjectToMoveObject.Add(objects1);

                //NXOpen.Body body1 = ((NXOpen.Body)workPart.Bodies.FindObject("UNPARAMETERIZED_FEATURE(2)"));
                //bool added1;
                //added1 = moveObjectBuilder1.ObjectToMoveObject.Add(body1);

                NXOpen.Point3d manipulatororigin1 = new NXOpen.Point3d(367.64618968964987, 6.6911544799500042, 10.0);
                moveObjectBuilder1.TransformMotion.TempManipulatorOrigin = manipulatororigin1;

                NXOpen.Point3d manipulatororigin2 = new NXOpen.Point3d(war, Convert.ToDouble(ws.Cells[i, 2].Value2.ToString()), Convert.ToDouble(ws.Cells[i, 3].Value2.ToString()));
                moveObjectBuilder1.TransformMotion.ManipulatorOrigin = manipulatororigin2;

                NXOpen.Matrix3x3 manipulatormatrix1 = new NXOpen.Matrix3x3();
                manipulatormatrix1.Xx = 1.0;
                manipulatormatrix1.Xy = 0.0;
                manipulatormatrix1.Xz = 0.0;
                manipulatormatrix1.Yx = 0.0;
                manipulatormatrix1.Yy = 1.0;
                manipulatormatrix1.Yz = 0.0;
                manipulatormatrix1.Zx = 0.0;
                manipulatormatrix1.Zy = 0.0;
                manipulatormatrix1.Zz = 1.0;
                moveObjectBuilder1.TransformMotion.ManipulatorMatrix = manipulatormatrix1;

                NXOpen.Session.UndoMarkId markId2;
                markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Move Object");

                theSession.DeleteUndoMark(markId2, null);

                NXOpen.Session.UndoMarkId markId3;
                markId3 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Move Object");

                NXOpen.NXObject nXObject1;
                nXObject1 = moveObjectBuilder1.Commit();

                //NXOpen.NXObject[] objects1;
                objects1 = moveObjectBuilder1.GetCommittedObjects();

                theSession.DeleteUndoMark(markId3, null);
                theSession.SetUndoMarkName(markId1, "Move Object");

                moveObjectBuilder1.Destroy();
            } 
        }
        wb.Close();
    }
    public static int GetUnloadOption(string dummy) { return (int)NXOpen.Session.LibraryUnloadOption.Immediately; }
}
