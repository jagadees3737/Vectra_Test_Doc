using System;
using INFITF;
using MECMOD;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DRAFTINGITF;

namespace Vectra_test
{
    /*The Following class is used to Create a Drawing view from a specific part body.
     * The name of the part should be "F497325052_A.CATPart".
     * The name of the Body to project a view should be as "Flame_Cut".
     * The view should be named as "Flame_Cut_View".
     * The View should be placed in the center of the A4 sized sheet with default drawing template.
     */
    class Scope_2
    {
        static INFITF.Application myCatia;

        /// <summary>
        /// To initiate or start the CATIA application.
        /// </summary>
        public void startCatia()
        {
            try
            {
                myCatia = (INFITF.Application)Marshal.GetActiveObject("CATIA.Application");
            }
            catch (Exception)
            {
                myCatia = (INFITF.Application)Activator.CreateInstance(Type.GetTypeFromProgID("CATIA.Application"));
            }
            Drawing_view();
        }

        /// <summary>
        /// Method for creating a drawing view.
        /// </summary>
        private void Drawing_view()
        {
            //To set up the Part Document
            PartDocument partDocument = (PartDocument)myCatia.Documents.Open("F497325052_A.CATPart");
            Part part1 = (Part)partDocument.Part;
            Bodies bodies = part1.Bodies;

            //To get the Body with the Specified name ("Flame_Cut").
            for (int bdy_count = 1; bdy_count <= bodies.Count; bdy_count++)
            {
                if (bodies.Item(bdy_count).get_Name() == "Flame_Cut")
                {
                    //To set up the drawing sheet with A4 size and Default Template.
                    DrawingDocument drawingDocument = (DrawingDocument)myCatia.Documents.Add("Drawing");
                    DrawingSheets sheets = (DrawingSheets)drawingDocument.Sheets;
                    DrawingSheet activeSheet = sheets.ActiveSheet;
                    activeSheet.set_Name("Flame_Cut_View");
                    activeSheet.PaperSize = DRAFTINGITF.CatPaperSize.catPaperA4;
                    int paperHeight = (int)activeSheet.GetPaperHeight();
                    int paperWidth = (int)activeSheet.GetPaperWidth();

                    //To Insert the front view of the specified body ("Flame_Cut").
                    DrawingViews views = activeSheet.Views;
                    DrawingView drawingView = (DrawingView)views.Add("FrontView");
                    drawingView.y = (double)(paperHeight / 2);
                    drawingView.x = (double)(paperWidth / 2);
                    drawingView.Scale = 1.0;
                    DrawingViewGenerativeLinks linkbehaviour = drawingView.GenerativeLinks;
                    DrawingViewGenerativeBehavior generativeBehavior = drawingView.GenerativeBehavior;
                    Body prtbody = (Body)bodies.GetItem("Flame_Cut");
                    linkbehaviour.AddLink(prtbody);
                    drawingDocument.Update();
                }

                //To Print message incase the Specified body is not present in the Document.
                else
                {
                    MessageBox.Show("The required Body is not present");
                }
            }
        }
        
    }
}
