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

            switch (document3D.DocumentType)
            {
                case DocumentTypeEnum.ksDocumentUnknown:
                    break;
                case DocumentTypeEnum.ksDocumentDrawing:
                    break;
                case DocumentTypeEnum.ksDocumentFragment:
                    break;
                case DocumentTypeEnum.ksDocumentSpecification:
                    break;
                case DocumentTypeEnum.ksDocumentPart:
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
                            }
                            if (item.Name == "Обозначение")
                            {
                                dynamic info;
                                bool source;
                                propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                                partDesignation = info;
                            }
                            if (item.Name == "Материал")
                            {
                                dynamic info;
                                bool source;
                                propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                                partMaterial = info;
                            }
                            if (item.Name == "Масса")
                            {
                                item.SignificantDigitsCount = 2;
                                dynamic info;
                                bool source;
                                propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                                partMass = info;
                            }
                            if (item.Name == "Раздел спецификации")
                            {
                                dynamic info;
                                bool source;
                                propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
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
                        
                        if (t!=null)
                        {
                            string NameProgramm = t.Value + "mm_" + partDesignation.Remove(0, 3);
                            worksheet.Cell(18, 3).Value = NameProgramm;
                            worksheet.Cell(18, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                            worksheet.Cell(11, 7).Value = t.Value;
                        }                        
                        #endregion
                        excelWorkbook.SaveAs(PathName + partDesignation + " - " + partName + ".xlsx");
                    }
                    break;
                case DocumentTypeEnum.ksDocumentAssembly:
                    {
                        string a = Path.Combine(Environment.CurrentDirectory, "PartTemplate.xlsx");
                        string PathName = document3D.Path;

                        KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
                        kompas.Visible = true;
                        kompas.ActivateControllerAPI();
                        ksDocument3D ksDocument3D = (ksDocument3D)kompas.ActiveDocument3D();
                        ksPartCollection _ksPartCollection = ksDocument3D.PartCollection(true);
                        Dictionary<string, int> collectionParts = new Dictionary<string, int>();
                        Dictionary<string, int> collectionStandartDetails = new Dictionary<string, int>();
                        Dictionary<string, int> othertDetails = new Dictionary<string, int>();
                        for (int i = 0; i < _ksPartCollection.GetCount(); i++)
                        {
                            ksPart ksPart = _ksPartCollection.GetByIndex(i);
                            IPart7 _part7 = kompas.TransferInterface(ksPart, 2, 0);

                            IApplication _application = kompas.ksGetApplication7();
                            IKompasDocument3D kompasDocument3D = (IKompasDocument3D)_application.ActiveDocument;
                            IPropertyMng propertyMng = (IPropertyMng)_application;
                            var properties = propertyMng.GetProperties(kompasDocument3D);
                            IPropertyKeeper propertyKeeper = (IPropertyKeeper)_part7;

                            string partName = "";
                            foreach (IProperty item in properties)
                            {
                                if (item.Name == "Раздел спецификации")
                                {
                                    dynamic info;
                                    bool source;
                                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                                    partName = info;
                                }
                            }
                            if (collectionParts.ContainsKey(ksPart.marking + " - " + ksPart.name)
                                && ksPart.marking != ""
                                && ksPart.excluded == false)
                            {
                                collectionParts[ksPart.marking + " - " + ksPart.name] = collectionParts[ksPart.marking + " - " + ksPart.name] + 1;
                            }
                            else if (collectionParts.ContainsKey(ksPart.marking + " - " + ksPart.name) == false
                                     && ksPart.marking != ""
                                     && ksPart.excluded == false)
                            { collectionParts.Add(ksPart.marking + " - " + ksPart.name, 1); }
                            else if (collectionStandartDetails.ContainsKey(ksPart.name)
                                //&& ksPart.marking == ""
                                && ksPart.excluded == false
                                && partName == "Стандартные изделия")
                            {
                                collectionStandartDetails[ksPart.name] = collectionStandartDetails[ksPart.name] + 1;
                            }
                            else if (collectionStandartDetails.ContainsKey(ksPart.name) == false
                                     //&& ksPart.marking == ""
                                     && ksPart.excluded == false
                                     && partName == "Стандартные изделия")
                            { collectionStandartDetails.Add(ksPart.name, 1); }
                            else if (othertDetails.ContainsKey(ksPart.name)
                                //&& ksPart.marking == ""
                                && ksPart.excluded == false
                                && partName == "Прочие изделия")
                            {
                                othertDetails[ksPart.name] = othertDetails[ksPart.name] + 1;
                            }
                            else if (othertDetails.ContainsKey(ksPart.name) == false
                                     //&& ksPart.marking == ""
                                     && ksPart.excluded == false
                                     && partName == "Прочие изделия")
                            { othertDetails.Add(ksPart.name, 1); }
                        }
                    }
                    break;
                case DocumentTypeEnum.ksDocumentTextual:
                    break;
                case DocumentTypeEnum.ksDocumentTechnologyAssembly:
                    break;
                default:
                    break;
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Подключаемся к компасу
            IApplication application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
            IKompasDocument3D document3D = (IKompasDocument3D)application.ActiveDocument;
            IPart7 part7 = document3D.TopPart;

            #region Присваиваю раздел спецификации
            IPropertyMng propertyMng = (IPropertyMng)application;
            var properties = propertyMng.GetProperties(document3D);
            IPropertyKeeper propertyKeeper = (IPropertyKeeper)part7;
            foreach (IProperty item in properties)
            {
                if (item.Name == "Раздел спецификации")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    string otherPart = @"'<property id=""SPCSection"" expression="""" fromSource=""false"" format=""{$sectionName}"">''''<property id=""sectionName"" value=""Прочие изделия"" type=""string"" />''''<property id=""sectionNumb"" value=""30"" type=""int"" />'";
                    string detal = @"'<property id=""SPCSection"" expression="""" fromSource=""false"" format=""{$sectionName}"">''''<property id=""sectionName"" value=""Детали"" type=""string"" />''''<property id=""sectionNumb"" value=""20"" type=""int"" />'";
                    string assembly = @"'<property id=""SPCSection"" expression="""" fromSource=""false"" format=""{$sectionName}"">''''<property id=""sectionName"" value=""Сборочные единицы"" type=""string"" />''''<property id=""sectionNumb"" value=""15"" type=""int"" />'";
                    propertyKeeper.SetComplexPropertyValue((_Property)item, detal);
                }
            }
            #endregion
                        
            #region Проверяю совпадает ли имя и обозначение с именем файла
            IKompasDocument kompasDocument = (IKompasDocument)application.ActiveDocument;
            if (kompasDocument.Name == "")
            {
                MessageBox.Show("Сохраните деталь");
            }
            else
            {
                string partName1 = "";
                string partDesignation1 = "";
                IPropertyMng propertyMng1 = (IPropertyMng)application;
                var properties1 = propertyMng1.GetProperties(document3D);
                IPropertyKeeper propertyKeeper1 = (IPropertyKeeper)part7;
                foreach (IProperty item in properties)
                {
                    if (item.Name == "Наименование")
                    {
                        dynamic info;
                        bool source;
                        propertyKeeper1.GetPropertyValue((_Property)item, out info, false, out source);
                        partName1 = info;
                    }
                    if (item.Name == "Обозначение")
                    {
                        dynamic info;
                        bool source;
                        propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                        partDesignation1 = info;
                    }
                }
                MessageBox.Show(kompasDocument.Name.Remove(kompasDocument.Name.Count() - 4) + "   |   " + "Имя документа\n" + partDesignation1 + " - " + partName1 + "   |   " + "Имя/обозначение");
            }
            #endregion

            #region Проверяю совпадает ли глобальная переменная толщина с толщиной в определении листового тела
            IFeature7 pFeat = (IFeature7)part7.Owner;
            Object[] featCol = pFeat.SubFeatures[0, false, false];
            ////https://forum.ascon.ru/index.php?topic=31251.msg249518#msg249518
            
            double t=0;
            
            IFeature7 _feature7 = (IFeature7)document3D.TopPart;
            var _t = _feature7.Variable[false, true, "SM_Thickness"];

            foreach (IFeature7 item in featCol)
            {
                if(item.Name.Contains("Листовое тело:"))
                {
                    for (int i = 0; i < item.VariablesCount[false,true]; i++)
                    {
                        if (item.Variable[false, true, i].ParameterNote == @"Толщина листового тела")
                        {
                            t = item.Variable[false, true, i].Value;
                        }
                    }
                }                
            }
            if (t != _t.Value) { MessageBox.Show("Толщина глобальной переменной и толщина листового тела не совпадают"); }
            #endregion
        }

        private void button5_Click(object sender, EventArgs e)
        {
            KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            kompas.Visible = true;
            kompas.ActivateControllerAPI();
            ksDocument3D ksDocument3D = (ksDocument3D)kompas.ActiveDocument3D();
            ksPartCollection _ksPartCollection = ksDocument3D.PartCollection(true);
            for (int i = 0; i < _ksPartCollection.GetCount(); i++)
            {

            }
        }
    }
}
