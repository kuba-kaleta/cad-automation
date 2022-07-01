using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

using NXOpen;
using NXOpen.UF;
using NXOpenUI;


namespace TableTest
{
    class NewPart //Create a part in NX open C# Tamil - dziala
    {  
        public static void Main(string[] args)
        {
            Session session = Session.GetSession();
            UFSession uFSession = UFSession.GetUFSession();

            Tag tag = Tag.Null;

            uFSession.Part.New("MyPart6", 1, out tag);

            ListingWindow listingWindow = session.ListingWindow;
            listingWindow.Open();

            listingWindow.WriteLine(tag.ToString());
            listingWindow.WriteLine(session.Parts.Work.Name);
        }

        public static int GetUnloadOption(string arg)
        {
            return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);
        }

    }
}
