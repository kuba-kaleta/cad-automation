using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using NXOpen;
using NXOpenUI;
using NXOpen.UF;

namespace TableTest
{
    class Expression //How to Update Expressions using NX Open - dziala
    {
        public static void Main(string[] args)
        {
            Session session = Session.GetSession();
            UFSession uFSession = UFSession.GetUFSession();
            UI ui = UI.GetUI();
            try
            {
                string exp;
                uFSession.Modl.AskExp("length", out exp);

                // ui.NXMessageBox.Show("Current Value", NXMessageBox.DialogType.Information, exp);

                if (exp != "")
                {
                    uFSession.Modl.EditExp("length=120");
                    ui.NXMessageBox.Show("Current Value", NXMessageBox.DialogType.Information, "Success");

                    uFSession.Modl.Update();
                }
            }

            catch (Exception exception)
            {
                ui.NXMessageBox.Show("Error", NXMessageBox.DialogType.Information, exception.Message);

            }
        }

        public static int GetUnloadOption(string arg)
        {
            return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);

        }

    }
}
