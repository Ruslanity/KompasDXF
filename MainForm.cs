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
using System.Diagnostics;
using System.Xml;
using PdfiumViewer;

namespace Multitool
{
    public partial class MainForm : Form
    {
        private static MainForm instance;
        private DxfViewerControl dxfViewer;
        private Settings settings;
        private PdfViewerControl pdfControl;


        public static MainForm GetInstance() //реализация Singleton
        {
            if (instance == null || instance.IsDisposed)
            {
                instance = new MainForm();
            }
            return instance;
        }

        public class SettingsData
        {
            public string textBox_DXF { get; set; }
            public string textBox_PDF { get; set; }
            //public string textBox_CUT_SPEED { get; set; }
        }

        public static string GetSettingsFilePath()
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Rusik Edition", "Multitool");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "Settings.xml");
        }

        public static string GetTemplateDir()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "Rusik Edition", "Multitool");
        }

        public MainForm()
        {
            TopMost = true;
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
        }

        private void СreateDXF_Click(object sender, EventArgs e)
        {
            string save_to_name2 = null;
            string ucdName = null;
            //if (eDrawForm != null)
            //{
            //    eDrawForm.Dispose();
            //}
            if (pdfControl != null)
            {
                tableLayoutPanel1.Controls.Remove(pdfControl);
                pdfControl.Dispose();
                pdfControl = null;
            }
            if (dxfViewer != null)
            {
                tableLayoutPanel1.Controls.Remove(dxfViewer);
                dxfViewer.Dispose();
                dxfViewer = null;
            }

            IApplication application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
            IKompasDocument3D document3D = (IKompasDocument3D)application.ActiveDocument;
            IPart7 part7 = document3D.TopPart;

            ISheetMetalContainer sheetMetalContainer = part7 as ISheetMetalContainer;
            ISheetMetalBodies sheetMetalBodies = sheetMetalContainer.SheetMetalBodies;
            ISheetMetalBody sheetMetalBody = sheetMetalBodies.SheetMetalBody[0];

            FileInfo fi = new FileInfo(part7.FileName);

            string save_to_name = fi.DirectoryName + "\\" +
                sheetMetalBody.Thickness.ToString(CultureInfo.CreateSpecificCulture("en-GB")) + "mm_" + part7.Marking.Remove(0, 3) + ".dxf";

            #region Создаю путь куда еще будет копироваться DXF

            string filePath = GetSettingsFilePath();

            if (System.IO.File.Exists(filePath))
            {
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(SettingsData));
                using (var reader = new System.IO.StreamReader(filePath))
                {
                    var loadedData = (SettingsData)serializer.Deserialize(reader);
                    if (!string.IsNullOrEmpty(loadedData.textBox_DXF))
                    {
                        if (Directory.Exists(loadedData.textBox_DXF))
                        {
                            string dxfBase = sheetMetalBody.Thickness.ToString(CultureInfo.CreateSpecificCulture("en-GB")) + "mm_" + part7.Marking.Remove(0, 3);
                            save_to_name2 = Path.Combine(loadedData.textBox_DXF, dxfBase + ".dxf");
                            ucdName = Path.Combine(loadedData.textBox_DXF, dxfBase + ".ucd");
                        }
                        else
                        {
                            MessageBox.Show("Папка для DXF не найдена:\n" + loadedData.textBox_DXF);
                        }
                    }
                }
            }
            #endregion

            #region Удаляю ucd
            if (System.IO.File.Exists(ucdName))
            {
                System.IO.File.Delete(ucdName);
            }
            #endregion

            KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");

            ksDocumentParam documentParam = (ksDocumentParam)kompas.GetParamStruct(35);
            documentParam.type = 1;
            documentParam.Init();
            ksDocument2D document2D = (ksDocument2D)kompas.Document2D();
            document2D.ksCreateDocument(documentParam);

            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)application.ActiveDocument;

            //Скрываем все сообщения системы -Да
            application.HideMessage = ksHideMessageEnum.ksHideMessageYes;

            IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
            IViews views = viewsAndLayersManager.Views;
            IView pView = views.Add(Kompas6Constants.LtViewType.vt_Arbitrary);

            IAssociationView pAssociationView = pView as IAssociationView;
            pAssociationView.SourceFileName = part7.FileName;

            //скрываю оси при создании dxf
            IAssociationViewElements associationViewElements = (IAssociationViewElements)pAssociationView;
            associationViewElements.CreateCircularCentres = false;
            associationViewElements.CreateLinearCentres = false;
            associationViewElements.CreateAxis = false;
            associationViewElements.CreateCentresMarkers = false;
            associationViewElements.ProjectAxis = false;
            associationViewElements.ProjectDesTexts = false;

            IEmbodimentsManager embodimentsManager = (IEmbodimentsManager)document3D;
            int indexPart = embodimentsManager.CurrentEmbodimentIndex;

            IEmbodimentsManager emb = (IEmbodimentsManager)pAssociationView;
            emb.SetCurrentEmbodiment(part7.Marking);

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
            //Скрываем все сообщения системы -Нет
            application.HideMessage = ksHideMessageEnum.ksShowMessage;
            document2D.ksSaveDocument(save_to_name);
            if (save_to_name2 != null)
                document2D.ksSaveDocument(save_to_name2);

            IKompasDocument kompasDocument = (IKompasDocument)application.ActiveDocument;
            kompasDocument.Close(DocumentCloseOptions.kdDoNotSaveChanges);
            kompas = null;
            application = null;

            OpenDxfInViewer(save_to_name);
            
        }


        private void СreatePDF_Click(object sender, EventArgs e)
        {
            IApplication application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
            IKompasDocument3D document3D = (IKompasDocument3D)application.ActiveDocument;
            IPart7 part7 = document3D.TopPart;

            ISheetMetalContainer sheetMetalContainer = part7 as ISheetMetalContainer;
            ISheetMetalBodies sheetMetalBodies = sheetMetalContainer.SheetMetalBodies;
            ISheetMetalBody sheetMetalBody = sheetMetalBodies.SheetMetalBody[0];

            string drawingPath = /*document3D.PathName +*/ part7.FileName.Remove(part7.FileName.Length - 4) + ".cdw";
            string drawingName = Path.GetFileName(drawingPath);
            string drawingName2 = /*document3D.PathName +*/ part7.Marking + " - " + part7.Name + ".cdw";
            string folderName = Path.GetDirectoryName(drawingPath);

            #region Поиск номера исполнения
            IEmbodimentsManager embodimentsManager = (IEmbodimentsManager)document3D;
            int indexPart = embodimentsManager.CurrentEmbodimentIndex;
            string basename = embodimentsManager.GetCurrentEmbodimentMarking(ksVariantMarkingTypeEnum.ksVMBaseMarking, false);
            //IEmbodimentsManager embodimentsManager = (IEmbodimentsManager)part7;
            //var embodiment = embodimentsManager.GetCurrentEmbodimentMarking(ksVariantMarkingTypeEnum.ksVMEmbodimentNumber, false);
            //расчет индекса
            string indexPartString = (indexPart - 1).ToString("D2");
            string drawingName3 = String.Empty;
            if (document3D.DocumentType == DocumentTypeEnum.ksDocumentPart && sheetMetalBody != null)
            {
                drawingName3 = basename + "-" + indexPartString + " - " + part7.Name + ".cdw";
            }
            else
            {
                drawingName3 = basename + "-" + indexPartString + " СБ" + " - " + part7.Name + ".cdw";
            }
            #endregion

            string[] fileNames =
            {
                drawingName2,
                drawingName3,
                drawingName
            };

            string[] fileEntries = Directory.GetFiles(document3D.Path);

            // Определяем путь к PDF рядом с моделью (не зависит от настроек)
            string filenamePDF;
            string pdfFileName;
            if (document3D.DocumentType == DocumentTypeEnum.ksDocumentPart && sheetMetalBody != null)
            {
                pdfFileName = sheetMetalBody.Thickness.ToString(CultureInfo.CreateSpecificCulture("en-GB")) + "mm_" + part7.Marking.Remove(0, 3) + ".pdf";
                filenamePDF = Path.Combine(Path.GetDirectoryName(part7.FileName), pdfFileName);
            }
            else
            {
                pdfFileName = part7.Marking + " - " + part7.Name + ".pdf";
                filenamePDF = Path.Combine(folderName, pdfFileName);
            }

            #region Путь куда копировать
            string copyfilenamePDF = String.Empty;
            string filePath = GetSettingsFilePath();
            if (System.IO.File.Exists(filePath))
            {
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(SettingsData));
                using (var reader = new System.IO.StreamReader(filePath))
                {
                    var loadedData = (SettingsData)serializer.Deserialize(reader);
                    if (!string.IsNullOrEmpty(loadedData.textBox_PDF))
                    {
                        if (Directory.Exists(loadedData.textBox_PDF))
                            copyfilenamePDF = Path.Combine(loadedData.textBox_PDF, pdfFileName);
                        else
                            MessageBox.Show("Папка для PDF не найдена:\n" + loadedData.textBox_PDF);
                    }
                }
            }
            #endregion

            

            bool fileExists = false;
            foreach (var fileName in fileNames)
            {
                string fullPath = Path.Combine(folderName, fileName);
                if (fileEntries.Contains(fullPath))
                {

                    application.HideMessage = ksHideMessageEnum.ksHideMessageNo;
                    Converter сonverter = application.Converter[@"C:\Program Files\ASCON\KOMPAS-3D v22\Bin\Pdf2d.dll"];
                    сonverter.Convert(fullPath, filenamePDF, 0, true);
                    if (!string.IsNullOrEmpty(copyfilenamePDF))
                        сonverter.Convert(fullPath, copyfilenamePDF, 0, true);
                    application.HideMessage = ksHideMessageEnum.ksShowMessage;

                    #region Тут я показываю PDF в контроле
                    if (dxfViewer != null)
                    {
                        tableLayoutPanel1.Controls.Remove(dxfViewer);
                        dxfViewer.Dispose();
                        dxfViewer = null;
                    }
                    if (pdfControl != null)
                    {
                        tableLayoutPanel1.Controls.Remove(pdfControl);
                        pdfControl.Dispose();
                    }
                    this.pdfControl = new PdfViewerControl();
                    tableLayoutPanel1.Controls.Add(pdfControl, 1, 0);
                    tableLayoutPanel1.SetRowSpan(pdfControl, 9);
                    pdfControl.Dock = DockStyle.Fill;
                    if (pdfControl != null)
                    {
                        pdfControl.LoadPdf(filenamePDF);
                        statusLabel.Text = Path.GetFileName(filenamePDF);
                    }
                    fileExists = true;
                    break;
                    #endregion
                }
            }
            if (!fileExists)
            {
                MessageBox.Show("В каталоге нет одноименного чертежа");
            }

        }

        private void СreateExcel_Click(object sender, EventArgs e)
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
                        string a = Path.Combine(GetTemplateDir(), "PartTemplate.xlsx");
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

                        if (t != null)
                        {
                            string NameProgramm = t.Value + "mm_" + partDesignation.Remove(0, 3);
                            worksheet.Cell(18, 3).Value = NameProgramm;
                            worksheet.Cell(18, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                            worksheet.Cell(11, 8).Value = t.Value;
                        }
                        #endregion
                        excelWorkbook.SaveAs(PathName + partDesignation + " - " + partName + ".xlsx");
                    }
                    break;
                case DocumentTypeEnum.ksDocumentAssembly:
                    {
                        string a = "";
                        if (comboBox1.Text == "Сварочный")
                        {
                            a = Path.Combine(GetTemplateDir(), "AssemblyTemplateWeld.xlsx");
                        }
                        if (comboBox1.Text == "Метизный")
                        {
                            a = Path.Combine(GetTemplateDir(), "AssemblyTemplate.xlsx");
                        }

                        string PathName = document3D.Path;
                        #region Вытаскиваем свойства сборки
                        string partName = "";
                        string partDesignation = "";
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
                        }
                        #endregion
                        XLWorkbook excelWorkbook = new XLWorkbook(a);
                        IXLWorksheet worksheet = excelWorkbook.Worksheet(1);

                        #region Обозначение сборки
                        worksheet.Cell(11, 1).Value = partDesignation;
                        worksheet.Cell(11, 1).Style.Font.FontName = "Arial Cyr";
                        worksheet.Cell(11, 1).Style.Font.Bold = false;
                        worksheet.Cell(11, 1).Style.Font.Italic = false;
                        worksheet.Cell(11, 1).Style.Font.FontSize = 11;
                        worksheet.Cell(11, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        worksheet.Cell(11, 1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                        #endregion

                        #region Наименование сборки
                        worksheet.Cell(11, 4).Value = partName;
                        worksheet.Cell(11, 4).Style.Font.FontName = "Arial Cyr";
                        worksheet.Cell(11, 4).Style.Font.Italic = false;
                        worksheet.Cell(11, 4).Style.Font.FontSize = 11;
                        worksheet.Cell(11, 4).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        worksheet.Cell(11, 4).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                        #endregion

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
                            IPropertyMng _propertyMng = (IPropertyMng)_application;
                            var _properties = _propertyMng.GetProperties(kompasDocument3D);
                            IPropertyKeeper _propertyKeeper = (IPropertyKeeper)_part7;

                            string partSection = "";
                            foreach (IProperty item in _properties)
                            {
                                if (item.Name == "Раздел спецификации")
                                {
                                    dynamic info;
                                    bool source;
                                    _propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                                    partSection = info;
                                }
                            }
                            #region Заполняю коллекции                            
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
                                && partSection == "Стандартные изделия")
                            {
                                collectionStandartDetails[ksPart.name] = collectionStandartDetails[ksPart.name] + 1;
                            }
                            else if (collectionStandartDetails.ContainsKey(ksPart.name) == false
                                     //&& ksPart.marking == ""
                                     && ksPart.excluded == false
                                     && partSection == "Стандартные изделия")
                            { collectionStandartDetails.Add(ksPart.name, 1); }
                            else if (othertDetails.ContainsKey(ksPart.name)
                                //&& ksPart.marking == ""
                                && ksPart.excluded == false
                                && partSection == "Прочие изделия")
                            {
                                othertDetails[ksPart.name] = othertDetails[ksPart.name] + 1;
                            }
                            else if (othertDetails.ContainsKey(ksPart.name) == false
                                     //&& ksPart.marking == ""
                                     && ksPart.excluded == false
                                     && partSection == "Прочие изделия")
                            { othertDetails.Add(ksPart.name, 1); }
                            #endregion
                        }

                        int quantityRows = Math.Max(collectionParts.Count, collectionStandartDetails.Count);

                        for (int i = 0; i < quantityRows; i++)
                        {
                            worksheet.Row(i + 15).InsertRowsBelow(1);
                            worksheet.Row(i + 15).Height = 30;
                        }

                        #region Задаю стиль ячеек входящих деталей
                        var myCustomStyle = XLWorkbook.DefaultStyle;
                        myCustomStyle.Font.FontName = "Arial Cyr";
                        myCustomStyle.Font.Bold = false;
                        myCustomStyle.Font.Italic = false;
                        myCustomStyle.Font.FontSize = 10;
                        myCustomStyle.Alignment.WrapText = true;
                        myCustomStyle.Border.LeftBorder = XLBorderStyleValues.Thin;
                        myCustomStyle.Border.RightBorder = XLBorderStyleValues.Thin;
                        myCustomStyle.Border.TopBorder = XLBorderStyleValues.Thin;
                        myCustomStyle.Border.BottomBorder = XLBorderStyleValues.Thin;
                        myCustomStyle.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                        myCustomStyle.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                        #endregion

                        #region Задаю стиль ячеек кол-ва деталей
                        var myCustomStyle2 = XLWorkbook.DefaultStyle;
                        myCustomStyle2.Font.FontName = "Arial Cyr";
                        myCustomStyle2.Font.Bold = false;
                        myCustomStyle2.Font.Italic = false;
                        myCustomStyle2.Font.FontSize = 10;
                        myCustomStyle2.Border.LeftBorder = XLBorderStyleValues.Thin;
                        myCustomStyle2.Border.RightBorder = XLBorderStyleValues.Thin;
                        myCustomStyle2.Border.TopBorder = XLBorderStyleValues.Thin;
                        myCustomStyle2.Border.BottomBorder = XLBorderStyleValues.Thin;
                        myCustomStyle2.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        myCustomStyle2.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                        #endregion

                        for (int i = 0; i < collectionParts.Count; i++)
                        {
                            IXLRange groop = worksheet.Range(String.Format("B{0}:D{1}", i + 15, i + 15)).Merge();
                            groop.Style = myCustomStyle;
                            groop.Value = collectionParts.ElementAt(i).Key;

                            worksheet.Cell(i + 15, 5).Value = collectionParts.ElementAt(i).Value;
                            worksheet.Cell(i + 15, 5).Style = myCustomStyle2;
                        }
                        if (collectionStandartDetails.Count != 0)
                        {
                            #region Шапка таблички
                            IXLRange groop = worksheet.Range("G13:I13").Merge();
                            groop.Value = "Метизы, входящие в сборку";
                            groop.Style.Font.FontName = "Arial Cyr";
                            groop.Style.Font.Bold = true;
                            groop.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            IXLRange groop2 = worksheet.Range("G14:H14").Merge();
                            groop2.Value = "№ Деталей";
                            groop2.Style = myCustomStyle2;
                            worksheet.Cell(14, 9).Value = "Кол-во";
                            worksheet.Cell(14, 9).Style = myCustomStyle2;
                            #endregion
                            for (int i = 0; i < collectionStandartDetails.Count; i++)
                            {
                                IXLRange groop3 = worksheet.Range(String.Format("G{0}:H{1}", i + 15, i + 15)).Merge();
                                groop3.Style = myCustomStyle;
                                groop3.Value = collectionStandartDetails.ElementAt(i).Key;
                                worksheet.Cell(i + 15, 9).Value = collectionStandartDetails.ElementAt(i).Value;
                                worksheet.Cell(i + 15, 9).Style = myCustomStyle2;
                            }
                        }
                        if (othertDetails.Count != 0)
                        {
                            worksheet.Row(16 + quantityRows).InsertRowsAbove(1);
                            worksheet.Row(16 + quantityRows).Height = 12.75;
                            worksheet.Row(16 + quantityRows).InsertRowsAbove(1);
                            worksheet.Row(16 + quantityRows).Height = 12.75;
                            for (int i = 0; i < othertDetails.Count; i++)
                            {
                                worksheet.Row(i + 18 + quantityRows).InsertRowsAbove(1);
                                worksheet.Row(i + 18 + quantityRows).Height = 30;
                            }
                            #region Шапка таблички
                            IXLRange groop = worksheet.Range(String.Format("B{0}:E{1}", quantityRows + 16, quantityRows + 16)).Merge();
                            groop.Value = "Прочие материалы:";
                            groop.Style.Font.FontName = "Arial Cyr";
                            groop.Style.Font.Bold = true;
                            groop.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            IXLRange groop2 = worksheet.Range(String.Format("B{0}:D{1}", quantityRows + 17, quantityRows + 17)).Merge();
                            groop2.Value = "№ Деталей";
                            groop2.Style = myCustomStyle2;
                            worksheet.Cell(quantityRows + 17, 5).Value = "Кол-во";
                            worksheet.Cell(quantityRows + 17, 5).Style = myCustomStyle2;
                            #endregion
                            for (int i = 0; i < othertDetails.Count; i++)
                            {
                                IXLRange groop3 = worksheet.Range(String.Format("B{0}:D{1}", i + 18 + quantityRows, i + 18 + quantityRows)).Merge();
                                groop3.Style = myCustomStyle;
                                groop3.Value = othertDetails.ElementAt(i).Key;
                                worksheet.Cell(i + 18 + quantityRows, 5).Value = othertDetails.ElementAt(i).Value;
                                worksheet.Cell(i + 18 + quantityRows, 5).Style = myCustomStyle2;
                            }
                        }
                        excelWorkbook.SaveAs(PathName + partDesignation + " - " + partName + ".xlsx");
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
                    if (part7.Detail == true)
                    {
                        propertyKeeper.SetComplexPropertyValue((_Property)item, detal);
                    }
                    else
                    {
                        propertyKeeper.SetComplexPropertyValue((_Property)item, assembly);
                    }
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

            double t = 0;

            IFeature7 _feature7 = (IFeature7)document3D.TopPart;
            var _t = _feature7.Variable[false, true, "SM_Thickness"];

            foreach (IFeature7 item in featCol)
            {
                if (item.Name.Contains("Листовое тело:"))
                {
                    for (int i = 0; i < item.VariablesCount[false, true]; i++)
                    {
                        if (item.Variable[false, true, i].ParameterNote == @"Толщина листового тела")
                        {
                            t = item.Variable[false, true, i].Value;
                        }
                    }
                }
            }
            if (_t != null)
            {
                if (t != _t.Value) { MessageBox.Show("Толщина глобальной переменной и толщина листового тела не совпадают"); }
            }

            #endregion
        }

        private void newDXF_Click(object sender, EventArgs e)
        {
            IApplication application;
            IKompasDocument3D document3D;
            IPart7 part7;
            KompasObject kompas;
            ksDocument3D doc3D;

            try
            {
                application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
                document3D = (IKompasDocument3D)application.ActiveDocument;
                part7 = document3D.TopPart;
                kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
                doc3D = (ksDocument3D)kompas.TransferInterface(document3D, 1, 0);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось подключиться к КОМПАС:\n" + ex.Message,
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (doc3D == null)
            {
                MessageBox.Show("Не удалось получить интерфейс ksDocument3D.",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var form = new FaceSelectForm(doc3D, document3D, part7, kompas, application);
            form.DxfCreated += ShowDxfInViewer;
            form.Show();
        }

        private void ShowDxfInViewer(string dxfPath)
        {
            if (pdfControl != null)
            {
                tableLayoutPanel1.Controls.Remove(pdfControl);
                pdfControl.Dispose();
                pdfControl = null;
            }
            OpenDxfInViewer(dxfPath);
        }

        private void OpenDxfInViewer(string dxfPath)
        {
            if (dxfViewer == null)
            {
                dxfViewer = new DxfViewerControl { Dock = DockStyle.Fill };
                tableLayoutPanel1.Controls.Add(dxfViewer, 1, 0);
                tableLayoutPanel1.SetRowSpan(dxfViewer, 9);
            }

            dxfViewer.LoadDxf(dxfPath);
            statusLabel.Text = Path.GetFileName(dxfPath);
        }

        private void Settings_Click(object sender, EventArgs e)
        {
            if (settings != null && !settings.IsDisposed)
            {
                // Если форма уже есть, активируем её
                settings.BringToFront();
                return;
            }
            settings = new Settings
            {
                TopMost = true
            };
            GroupBox groupBox = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "Пока не придумал"
            };
            settings.Controls.Add(groupBox);
            TableLayoutPanel tablelayoutpanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 5
            };
            tablelayoutpanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            tablelayoutpanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            groupBox.Controls.Add(tablelayoutpanel);

            Label label_DXF = new Label
            {
                Dock = DockStyle.Fill,
                Text = "Куда выгрузить DXF",
                TextAlign = ContentAlignment.MiddleLeft
            };
            tablelayoutpanel.Controls.Add(label_DXF, 0, 0);

            Label label_PDF = new Label
            {
                Dock = DockStyle.Fill,
                Text = "Куда выгрузить PDF",
                TextAlign = ContentAlignment.MiddleLeft
            };
            tablelayoutpanel.Controls.Add(label_PDF, 0, 1);

            //Label label_CUT_SPEED = new Label
            //{
            //    Dock = DockStyle.Fill,
            //    Text = "Путь до настроек скорости резки",
            //    TextAlign = ContentAlignment.MiddleLeft
            //};
            //tablelayoutpanel.Controls.Add(label_CUT_SPEED, 0, 1);

            TextBox textBox_DXF = new TextBox
            {
                Dock = DockStyle.Fill,
                TextAlign = HorizontalAlignment.Left
            };
            tablelayoutpanel.Controls.Add(textBox_DXF, 1, 0);

            TextBox textBox_PDF = new TextBox
            {
                Dock = DockStyle.Fill,
                TextAlign = HorizontalAlignment.Left
            };
            tablelayoutpanel.Controls.Add(textBox_PDF, 1, 1);

            //TextBox textBox_CUT_SPEED = new TextBox
            //{
            //    Dock = DockStyle.Fill,
            //    TextAlign = HorizontalAlignment.Left
            //};
            //tablelayoutpanel.Controls.Add(textBox_CUT_SPEED, 1, 1);


            string filePath = GetSettingsFilePath();

            //Дисериализация настроек
            if (System.IO.File.Exists(filePath))
            {
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(SettingsData));
                using (var reader = new System.IO.StreamReader(filePath))
                {
                    var loadedData = (SettingsData)serializer.Deserialize(reader);
                    textBox_DXF.Text = loadedData.textBox_DXF;
                    textBox_PDF.Text = loadedData.textBox_PDF;
                    //textBox_CUT_SPEED.Text = loadedData.textBox_CUT_SPEED;
                }
            }

            //Создаем объект для сериализации с текущим значением TextBox
            settings.FormClosing += (s, args) =>
            {
                var dataToSerialize = new SettingsData
                {
                    textBox_DXF = textBox_DXF.Text,
                    textBox_PDF = textBox_PDF.Text,
                    //textBox_CUT_SPEED = textBox_CUT_SPEED.Text
                };
                Settings_FormClosing(dataToSerialize, filePath);
            };
            settings.Show();
        }

        private void Settings_FormClosing(SettingsData data, string filePath)
        {
            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(SettingsData));
            using (var writer = new System.IO.StreamWriter(filePath))
            {
                serializer.Serialize(writer, data);
            }
        }
    }
}
