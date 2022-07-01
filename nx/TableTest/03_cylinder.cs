using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

using NXOpen;
using NXOpen.UF;
using NXOpenUI;
using NXOpen.Utilities;

namespace TableTest
{
    class Cylinder //How to Create a Cylinder on a selected Face in NX Open C# - nie dziala (Dialog area 2 is not currently available)
    {
        Session session = Session.GetSession();
        static UFSession uFSession = UFSession.GetUFSession();

        public static void Main(string[] args)
        {
            UFUi.SelInitFnT selInitFnT = new UFUi.SelInitFnT(selectionParams);
            int scope = UFConstants.UF_UI_SEL_SCOPE_WORK_PART;
            IntPtr intPtr = IntPtr.Zero;

            int outPut;
            Tag tag;
            Tag viewtag;
            double[] position = new double[3];
            uFSession.Ui.SelectWithSingleDialog("Single Dialog Test", "Select Any Face", scope, selInitFnT, intPtr, out outPut, out tag, position, out viewtag);
        }

        static int selectionParams(IntPtr select, IntPtr userData)
        {
            int num_triples = 1;
            UFUi.Mask[] masks = new UFUi.Mask[1];
            masks[0].object_type = UFConstants.UF_face_type;
            masks[0].object_subtype = 0;
            masks[0].solid_type = 0;

            uFSession.Ui.SetSelMask(select, UFUi.SelMaskAction.SelMaskClearAndEnableSpecific, num_triples, masks);
            return UFConstants.UF_UI_SEL_SUCCESS;

        }

        public static int GetUnloadOption(string arg)
        {
            return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);
        }

    }
}
