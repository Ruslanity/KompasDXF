using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;

namespace KompasDXF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            IApplication application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
            IKompasDocument3D document3D = (IKompasDocument3D)application.ActiveDocument;
            IPart7 part7 = document3D.TopPart;

            ISheetMetalContainer sheetMetalContainer = part7 as ISheetMetalContainer;
            ISheetMetalBodies sheetMetalBodies = sheetMetalContainer.SheetMetalBodies;
            ISheetMetalBody sheetMetalBody = sheetMetalBodies.SheetMetalBody[0];

            FileInfo fi = new FileInfo(part7.FileName);

            string save_to_name = fi.DirectoryName + "\\" +
                sheetMetalBody.Thickness.ToString(CultureInfo.CreateSpecificCulture("en-GB")) + "mm_"+ part7.Marking.Remove(0,3) + ".dxf";

            KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");

            ksDocumentParam documentParam = (ksDocumentParam)kompas.GetParamStruct(35);
            documentParam.type = 1;
            documentParam.Init();
            ksDocument2D document2D = (ksDocument2D)kompas.Document2D();
            document2D.ksCreateDocument(documentParam);

            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)application.ActiveDocument;

            //Скрываем все сообщения системы - Да
            application.HideMessage = ksHideMessageEnum.ksHideMessageYes;

            IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
            IViews views = viewsAndLayersManager.Views;
            IView pView = views.Add(Kompas6Constants.LtViewType.vt_Arbitrary);

            IAssociationView pAssociationView = pView as IAssociationView;
            pAssociationView.SourceFileName = part7.FileName;

            IEmbodimentsManager embodimentsManager = (IEmbodimentsManager)document3D;
            int indexPart = embodimentsManager.CurrentEmbodimentIndex;

            IEmbodimentsManager emb = (IEmbodimentsManager)pAssociationView;
            emb.SetCurrentEmbodiment(indexPart);

            pAssociationView.Angle = 0;
            pAssociationView.X = 0;
            pAssociationView.Y = 0;
            pAssociationView.BendLinesVisible = false;
            pAssociationView.BreakLinesVisible = false;
            pAssociationView.HiddenLinesVisible = false;
            pAssociationView.VisibleLinesStyle = (int)ksCurveStyleEnum.ksCSNormal;
            pAssociationView.Scale = 1;
            pAssociationView.Name = "User view";
            pAssociationView.ProjectionName = "#Развертка";
            pAssociationView.Unfold = true; //развернутый вид
            pAssociationView.BendLinesVisible = false;
            pAssociationView.CenterLinesVisible = false;
            pAssociationView.SourceFileName = part7.FileName;
            pAssociationView.Update();
            pView.Update();

            IViewDesignation pViewDesignation = pView as IViewDesignation;
            pViewDesignation.ShowUnfold = false;
            pViewDesignation.ShowScale = false;
                        
            pView.Update();
            document2D.ksRebuildDocument();
            //Скрываем все сообщения системы - Нет
            application.HideMessage = ksHideMessageEnum.ksHideMessageNo;
            document2D.ksSaveDocument(save_to_name);

            IKompasDocument kompasDocument = (IKompasDocument)application.ActiveDocument;
            kompasDocument.Close(DocumentCloseOptions.kdDoNotSaveChanges);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            IApplication application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
            IKompasDocument3D document3D = (IKompasDocument3D)application.ActiveDocument;
            IPart7 part7 = document3D.TopPart;

            string drawingName = /*document3D.PathName +*/ part7.FileName.Remove(part7.FileName.Length - 4) + ".cdw";
            string[] fileEntries = Directory.GetFiles(document3D.Path);
            if (fileEntries.Contains(drawingName))
            {                
                //Скрываем все сообщения системы - Да
                application.HideMessage = ksHideMessageEnum.ksHideMessageYes;
                //IKompasDocument2D kDoc = (IKompasDocument2D)application.Documents.Open(drawingName, true, false);
                //IKompasDocument2D1 kompasDocument2D1 = (IKompasDocument2D1)kDoc;
                //kompasDocument2D1.RebuildDocument();
                Converter сonverter = application.Converter[@"C:\\Program Files\\ASCON\\KOMPAS-3D v18\\Bin\Pdf2d.dll"];
                сonverter.Convert(part7.FileName.Remove(part7.FileName.Length - 4) + ".cdw",
                    part7.FileName.Remove(part7.FileName.Length - 4) + ".pdf",0,false);
                

                //Скрываем все сообщения системы - Нет
                application.HideMessage = ksHideMessageEnum.ksHideMessageNo;
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string workDirectory = Directory.GetCurrentDirectory();

            IApplication application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
            IKompasDocument3D document3D = (IKompasDocument3D)application.ActiveDocument;
            IPart7 part7 = document3D.TopPart;
            if (document3D.DocumentType == Kompas6Constants.DocumentTypeEnum.ksDocumentPart)
            {

            }

                MessageBox.Show(workDirectory);
        }
    }
}
