using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;
using DocumentFormat.OpenXml;

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
            IApplication application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
            IKompasDocument3D document3D = (IKompasDocument3D)application.ActiveDocument;
            IPart7 part7 = document3D.TopPart;
            
            if (document3D.DocumentType == Kompas6Constants.DocumentTypeEnum.ksDocumentPart)
            {
                string a = Path.Combine(Environment.CurrentDirectory, "PartTemplate.xlsx");
                string PathName = document3D.Path;

                #region Вытаскиваем свойства
                string partName = "";
                string partDesignation = "";
                string partMaterial = "";
                double partMass = 0;
                IPropertyMng propertyMng = (IPropertyMng)application;
                var properties = propertyMng.GetProperties(document3D);
                IPropertyKeeper propertyKeeper = (IPropertyKeeper)part7;
                foreach (IProperty item in properties)
                {
                    if (item.Name == "Наименование")
                    {
                        dynamic info;
                        bool source;
                        propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                        partName = info;
                        //SqlCommand cmd3 = new SqlCommand($"UPDATE [Detail] SET [FileName] = N'{info}' WHERE [Name] = N'{part.Name}'", SqlConnection);
                        //cmd3.ExecuteNonQuery();
                    }
                    if (item.Name == "Обозначение")
                    {
                        dynamic info;
                        bool source;
                        propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                        partDesignation = info;
                        //SqlCommand cmd3 = new SqlCommand($"UPDATE [Detail] SET [Designation] = N'{info}' WHERE [Name] = N'{part.Name}'", SqlConnection);
                        //cmd3.ExecuteNonQuery();
                    }
                    if (item.Name == "Материал")
                    {
                        dynamic info;
                        bool source;
                        propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                        partMaterial = info;
                        //SqlCommand cmd3 = new SqlCommand($"UPDATE [Detail] SET [Material] = N'{info}' WHERE [Name] = N'{part.Name}'", SqlConnection);
                        //cmd3.ExecuteNonQuery();
                    }
                    if (item.Name == "Масса")
                    {
                        item.SignificantDigitsCount = 2;
                        dynamic info;
                        bool source;
                        propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                        partMass = info;
                        //SqlCommand cmd3 = new SqlCommand($"UPDATE [Detail] SET [Mass] = N'{info}' WHERE [Name] = N'{part.Name}'", SqlConnection);
                        //cmd3.ExecuteNonQuery();
                    }
                    if (item.Name == "Раздел спецификации")
                    {
                        dynamic info;
                        bool source;
                        propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                        //SqlCommand cmd3 = new SqlCommand($"UPDATE [Detail] SET [Раздел спецификации] = N'{info}' WHERE [Name] = N'{part.Name}'", SqlConnection);
                        //cmd3.ExecuteNonQuery();
                    }
                }

                #endregion


                XLWorkbook excelWorkbook = new XLWorkbook(a);
                IXLWorksheet worksheet = excelWorkbook.Worksheet(1);
                #region Обозначение
                worksheet.Cell(11, 1).Value = partDesignation;
                worksheet.Cell(11, 1).Style.Font.FontName = "Arial Cyr";
                worksheet.Cell(11, 1).Style.Font.Bold = false;
                worksheet.Cell(11, 1).Style.Font.Italic = false;
                worksheet.Cell(11, 1).Style.Font.FontSize = 11;
                worksheet.Cell(11, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Cell(11, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                #endregion

                #region Наименование
                worksheet.Cell(11, 4).Value = partName;
                worksheet.Cell(11, 4).Style.Font.FontName = "Arial Cyr";
                worksheet.Cell(11, 4).Style.Font.Italic = false;
                worksheet.Cell(11, 4).Style.Font.FontSize = 11;
                worksheet.Cell(11, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Cell(11, 4).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                #endregion

                #region Материал
                worksheet.Cell(11, 6).Value = partMaterial;
                worksheet.Cell(11, 6).Style.Font.FontName = "Arial Cyr";
                worksheet.Cell(11, 1).Style.Font.Bold = false;
                worksheet.Cell(11, 6).Style.Font.Italic = false;
                worksheet.Cell(11, 6).Style.Font.FontSize = 10;                
                worksheet.Cell(11, 6).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Cell(11, 6).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                worksheet.Cell(11, 6).Style.Alignment.WrapText = true;
                #endregion

                #region Программа и толщина
                IFeature7 feature7 = (IFeature7)document3D.TopPart;
                var t = feature7.Variable[false, true, "SM_Thickness"];
                string NameProgramm = t.Value + "mm_" + partDesignation.Remove(0, 3);
                worksheet.Cell(18, 3).Value = NameProgramm;
                worksheet.Cell(18, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                worksheet.Cell(11, 7).Value = t.Value;

                #endregion


                excelWorkbook.SaveAs(PathName + partDesignation+" - "+ partName + ".xlsx");
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //KompasObject application = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            //if (application!=null)
            //{
            //    ksSpcDocument iDocumentSpc = (ksSpcDocument)application.SpcDocument();
            //    application.ActivateControllerAPI();
            //}
            IApplication application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
            IKompasDocument document = (IKompasDocument3D)application.ActiveDocument;
            ISpecificationDescriptions specDesc = document.SpecificationDescriptions;
            ISpecificationDescription specificationDescription = specDesc.ActiveFromLibStyle;

            if (specificationDescription == null)
            {
                specificationDescription = specDesc.Add(@"C:\Program Files\ASCON\KOMPAS-3D v18\Sys\graphic.lyt", 1, null);
                specificationDescription.DelegateMode = true;
                ISpecificationBaseObject specificationBaseObject = specificationDescription.BaseObjects.Add(20, 0);
                specificationBaseObject.SyncronizeWithProperties = true;
                specificationBaseObject.EditSourceObject = true;
                specificationBaseObject.Draw = true;
                specificationBaseObject.SpcUsed[0] = true;
                specificationBaseObject.Update();
                specificationDescription.Update();
                ISpecificationObject specificationObject = specificationBaseObject;
                //specificationObject.Edit();
                specificationObject.Update();
                //int s = specDesc.Count;
                //MessageBox.Show(s.ToString());
                //ISpecificationObject specificationObject = specificationDescription.Objects;
                //specificationObject.Update();
                //ISpecificationDocument specificationDocument = (ISpecificationDocument)document;
                //AttachedDocuments attachedDocuments = 
            }


            //specBaseObj[0].SetSection(20);
            //foreach (ISpecificationBaseObject item in specBaseObj)
            //{
            //    MessageBox.Show(item.Section.ToString());
            //}
            //ISpecificationBaseObject specificationBaseObject = specBaseObj[0];

        }
    }
}
