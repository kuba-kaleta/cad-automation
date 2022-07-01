using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using NXOpen;
using NXOpen.UF;

namespace TableTest
{
    class Test
    {
        public static void Main(string[] args)
        {
            NXOpen.Session theSession = NXOpen.Session.GetSession();
            NXOpen.Part workPart = theSession.Parts.Work;
            NXOpen.Part displayPart = theSession.Parts.Display;
            // ----------------------------------------------
            //   Menu: Tools->Journal->Stop Recording
            // ----------------------------------------------

            // ----------------------------------------------
            //   TODO: Message DONE
            // ----------------------------------------------

            UI theUI = UI.GetUI();
            int czy;

            czy = theUI.NXMessageBox.Show("Informancje", NXMessageBox.DialogType.Question, "Hello world");
            theUI.NXMessageBox.Show("Informancje", NXMessageBox.DialogType.Information, czy.ToString());
        }
    }
}