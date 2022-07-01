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
    class Cycle //Cycle Objects in Part Using NX Open - dziala
    {
        static Session session = Session.GetSession();
        static UFSession ufSession = UFSession.GetUFSession();

        public static void Main(string[] args)
        {
            Part part = session.Parts.Work;
            Tag objTag = Tag.Null;
            ListingWindow listingWindow = session.ListingWindow;
            listingWindow.Open();
            string names = "";
            do
            {
                objTag = ufSession.Obj.CycleAll(part.Tag, objTag);
                if (objTag != Tag.Null)
                {
                    TaggedObject taggedObject = NXObjectManager.Get(objTag);
                    names = names + taggedObject.ToString() + "\n";
                }
            } while (objTag != Tag.Null);

            listingWindow.WriteLine(names);
        }

        public static int GetUnloadOption(string arg)
        {
            return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);
        }
    }
}
