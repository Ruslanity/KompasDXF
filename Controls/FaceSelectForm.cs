using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using Kompas6API5;
using Kompas6Constants;
using Kompas6Constants3D;
using KompasAPI7;
namespace Multitool
{
    public class KompasViewRotator
    {
        private readonly IViewProjectionManager _viewProjMgr;

        public KompasViewRotator(IViewProjectionManager viewProjMgr)
        {
            _viewProjMgr = viewProjMgr;
        }

        /// <summary>
        /// Ориентирует 3D-вид так, чтобы нормаль грани смотрела в экран.
        /// Использует API7 IViewProjectionManager.SetMatrix3D.
        /// </summary>
        public bool OrientViewToFaceNormal(IFace face)
        {
            IMathSurface3D surface = face?.MathSurface;
            if (surface == null) return false;

            double u = (surface.ParamUMin + surface.ParamUMax) / 2.0;
            double v = (surface.ParamVMin + surface.ParamVMax) / 2.0;

            double nx, ny, nz;
            surface.GetNormal(u, v, out nx, out ny, out nz);

            // Учитываем флаг ориентации нормали грани
            if (!face.NormalOrientation) { nx = -nx; ny = -ny; nz = -nz; }

            double len = Math.Sqrt(nx * nx + ny * ny + nz * nz);
            if (len < 1e-9) return false;
            nx /= len; ny /= len; nz /= len;

            // up = мировой Z; если коллинеарен с нормалью — мировой Y
            double upx = 0, upy = 0, upz = 1;
            if (Math.Abs(nz) > 0.99) { upy = 1; upz = 0; }

            // правый вектор X = up × normal
            double vXx = upy * nz - upz * ny;
            double vXy = upz * nx - upx * nz;
            double vXz = upx * ny - upy * nx;
            double xLen = Math.Sqrt(vXx * vXx + vXy * vXy + vXz * vXz);
            if (xLen < 1e-9) return false;
            vXx /= xLen; vXy /= xLen; vXz /= xLen;

            // истинный вектор "вверх" Y = normal × X
            double vYx = ny * vXz - nz * vXy;
            double vYy = nz * vXx - nx * vXz;
            double vYz = nx * vXy - ny * vXx;

            // Матрица 3×4 (row-major):
            // Строка 0 = направление проекции (нормаль грани = ось глубины)
            // Строка 1 = горизонталь в плоскости проекции
            // Строка 2 = вертикаль в плоскости проекции
            double[] m = new double[16]
            {
                vXx, vXy, vXz, 0.0,
                vYx, vYy, vYz, 0.0,
                nx,  ny,  nz,  0.0,
                0, 0, 0, 0
            };

            // Применить ориентацию к живому 3D-окну
            _viewProjMgr.SetMatrix3D(m, 2.0);
            return true;
        }
    }

    public class FaceSelectForm : Form
    {
        private readonly ksDocument3D _doc3D;
        private readonly IKompasDocument3D _document3D;
        private readonly IPart7 _part7;
        private readonly KompasObject _kompas;
        private readonly IApplication _application;
        private ksSelectionMng _selMng;
        private IConnectionPoint _connPoint;
        private int _notifyCookie;
        private TextBox _faceTextBox;
        private object _selectedFaceObj;

        public event Action<string> DxfCreated;

        public FaceSelectForm(ksDocument3D doc3D, IKompasDocument3D document3D,
                              IPart7 part7, KompasObject kompas, IApplication application)
        {
            _doc3D = doc3D;
            _document3D = document3D;
            _part7 = part7;
            _kompas = kompas;
            _application = application;

            InitializeFormControls();
            SubscribeToSelectionEvents();

            FormClosing += (s, e) => UnsubscribeFromSelectionEvents();
        }

        private void InitializeFormControls()
        {
            Text = "Выбор грани";
            TopMost = true;
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            StartPosition = FormStartPosition.CenterScreen;
            Size = new System.Drawing.Size(340, 110);

            _faceTextBox = new TextBox
            {
                ReadOnly = true,
                Dock = DockStyle.Top,
                BackColor = System.Drawing.SystemColors.Control,
                Text = "Выберите грань в модели..."
            };
            Controls.Add(_faceTextBox);

            var btnPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                Padding = new Padding(5)
            };
            var btnOk = new Button { Text = "Ок" };
            var btnReset = new Button { Text = "Сброс" };
            btnOk.Click += BtnOk_Click;
            btnReset.Click += BtnReset_Click;
            btnPanel.Controls.Add(btnOk);
            btnPanel.Controls.Add(btnReset);
            Controls.Add(btnPanel);
        }

        private void SubscribeToSelectionEvents()
        {
            try
            {
                _selMng = (ksSelectionMng)_doc3D.GetSelectionMng();
                if (_selMng == null) { _faceTextBox.Text = "Ошибка: GetSelectionMng() вернул null"; return; }

                IConnectionPointContainer container = _selMng as IConnectionPointContainer;
                if (container == null) { _faceTextBox.Text = "Ошибка: нет IConnectionPointContainer"; return; }

                Guid guid = typeof(ksSelectionMngNotify).GUID;
                container.FindConnectionPoint(ref guid, out _connPoint);
                if (_connPoint == null) { _faceTextBox.Text = "Ошибка: FindConnectionPoint вернул null"; return; }

                _connPoint.Advise(new SelectionSink(this, _selMng), out _notifyCookie);
            }
            catch (Exception ex)
            {
                _faceTextBox.Text = "Ошибка подписки: " + ex.Message;
            }
        }

        private void UnsubscribeFromSelectionEvents()
        {
            try
            {
                if (_connPoint != null && _notifyCookie != 0)
                {
                    _connPoint.Unadvise(_notifyCookie);
                    _connPoint = null;
                    _notifyCookie = 0;
                }
            }
            catch { }
        }

        internal void OnFaceSelected(object faceObj)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action<object>(OnFaceSelected), faceObj); return; }
            _selectedFaceObj = faceObj;
            _faceTextBox.Text = "Грань выбрана";
        }

        internal void OnWrongTypeSelected(int objType)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action<int>(OnWrongTypeSelected), objType); return; }
            _faceTextBox.Text = "Тип объекта: " + objType + " (нужно 6 — грань)";
        }

        internal void OnSelectionCleared()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(OnSelectionCleared)); return; }
            _selectedFaceObj = null;
            _faceTextBox.Text = "Выберите грань в модели...";
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            try { _selMng?.UnselectAll(); } catch { }
            _selectedFaceObj = null;
            _faceTextBox.Text = "Выберите грань в модели...";
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (_selectedFaceObj == null)
            {
                MessageBox.Show("Сначала выберите грань в модели.", "Нет выбора",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                ISheetMetalContainer container = _part7 as ISheetMetalContainer;
                ISheetMetalBendUnfoldParameters unfoldParams = container.SheetMetalBendUnfoldParameters;

                // _selectedFaceObj — это API5-объект ksEntity из события выделения.
                // IModelObject (API7) нельзя получить прямым cast; нужен TransferInterface с ksAPI7Dual=2.
                IModelObject faceModelObj = _kompas.TransferInterface(_selectedFaceObj, 2, (int)Obj3dType.o3d_face) as IModelObject;
                if (faceModelObj == null)
                    throw new InvalidOperationException("Не удалось получить API7-интерфейс грани. Убедитесь, что выбрана грань детали из листового тела.");

                unfoldParams.DeleteParam();
                unfoldParams.UpdateParam();
                unfoldParams.FixedFaces = faceModelObj;
                unfoldParams.UnfoldPlane = faceModelObj;
                unfoldParams.UpdateParam();

                IFace face = (IFace)faceModelObj;
                IKompasDocument3D1 kompasDocument3D1 = (IKompasDocument3D1)_document3D;
                IViewProjectionManager viewProjectionManager = kompasDocument3D1.ViewProjectionManager;

                _document3D.SelectionManager.Select(face);

                // Удалить старые проекции "Авторазвертка" (при повторном запуске не накапливаются)
                for (int i = viewProjectionManager.Count - 1; i >= 0; i--)
                {
                    IViewProjection7 existingVp = viewProjectionManager.ViewProjection[i];
                    if (existingVp.Name == "Авторазвертка")
                        existingVp.Delete();
                }

                // Ориентировать 3D-вид нормально к выбранной грани через API7
                KompasViewRotator rotator = new KompasViewRotator(viewProjectionManager);
                if (!rotator.OrientViewToFaceNormal(face))
                    throw new InvalidOperationException("Не удалось ориентировать вид по нормали грани.");

                _part7.Update();

                // Захватить текущую ориентацию как именованную проекцию
                IViewProjection7 viewProjection7 = viewProjectionManager.Add();
                viewProjection7.Name = "Авторазвертка";
                viewProjection7.Update();
                _part7.Update();


                // Сохраняем модель на диск, чтобы ассоциативный вид при генерации DXF
                // читал файл с уже обновлёнными параметрами развёртки.
                _document3D.Save();
                //IKompasDocument modelDoc = (IKompasDocument)_document3D;
                //modelDoc.Save();

                string dxfPath = BuildDxfPath();
                CreateAndSaveDxf(dxfPath);

                DxfCreated?.Invoke(dxfPath);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при создании DXF:\n" + ex.Message,
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetDxfBaseName(ISheetMetalBody sheetMetalBody)
        {
            string marking = _part7.Marking ?? "";
            // Убираем префикс "АЛ." если есть
            if (marking.StartsWith("АЛ.", StringComparison.OrdinalIgnoreCase))
                marking = marking.Substring(3);
            // Если обозначение пустое — используем наименование
            if (string.IsNullOrWhiteSpace(marking))
                marking = _part7.Name ?? "";
            string thickness = sheetMetalBody.Thickness.ToString(CultureInfo.CreateSpecificCulture("en-GB"));
            return thickness + "mm_" + marking;
        }

        private string BuildDxfPath()
        {
            ISheetMetalContainer sheetMetalContainer = _part7 as ISheetMetalContainer;
            ISheetMetalBodies sheetMetalBodies = sheetMetalContainer.SheetMetalBodies;
            ISheetMetalBody sheetMetalBody = sheetMetalBodies.SheetMetalBody[0];

            FileInfo fi = new FileInfo(_part7.FileName);
            return fi.DirectoryName + "\\" + GetDxfBaseName(sheetMetalBody) + ".dxf";
        }

        private void CreateAndSaveDxf(string savePath)
        {
            string savePath2 = null;
            string ucdPath = null;

            ISheetMetalContainer sheetMetalContainer = _part7 as ISheetMetalContainer;
            ISheetMetalBodies sheetMetalBodies = sheetMetalContainer.SheetMetalBodies;
            ISheetMetalBody sheetMetalBody = sheetMetalBodies.SheetMetalBody[0];

            string filePath = MainForm.GetSettingsFilePath();

            if (File.Exists(filePath))
            {
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(MainForm.SettingsData));
                using (var reader = new StreamReader(filePath))
                {
                    var loadedData = (MainForm.SettingsData)serializer.Deserialize(reader);
                    if (!string.IsNullOrWhiteSpace(loadedData.textBox_DXF) && Directory.Exists(loadedData.textBox_DXF))
                    {
                        string baseName = GetDxfBaseName(sheetMetalBody);
                        savePath2 = loadedData.textBox_DXF + "\\" + baseName + ".dxf";
                        ucdPath = loadedData.textBox_DXF + "\\" + baseName + ".ucd";
                    }
                }
            }

            if (File.Exists(ucdPath))
                File.Delete(ucdPath);

            ksDocumentParam documentParam = (ksDocumentParam)_kompas.GetParamStruct(35);
            documentParam.type = 1;
            documentParam.Init();
            ksDocument2D document2D = (ksDocument2D)_kompas.Document2D();
            document2D.ksCreateDocument(documentParam);


            // Очистить технические требования: style=1 применяет пустую замену
            ksTechnicalDemandParam technicalDemandParam = (ksTechnicalDemandParam)_kompas.GetParamStruct(78);
            technicalDemandParam.Init();
            int TT = document2D.ksGetReferenceDocumentPart(1);
            if (TT != 0)
                document2D.ksDeleteObj(TT);

            // Удалить знак неуказанной шероховатости
            ksSpecRoughParam specRoughParam = (ksSpecRoughParam)_kompas.GetParamStruct(79);
            specRoughParam.Init();
            specRoughParam.sign = 0;
            int signRef = (int)document2D.ksSpecRough(specRoughParam);
            if (signRef != 0)
                document2D.ksDeleteObj(signRef);


            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)_application.ActiveDocument;
            _application.HideMessage = ksHideMessageEnum.ksHideMessageYes;

            IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
            IViews views = viewsAndLayersManager.Views;
            IView pView = views.Add(LtViewType.vt_Arbitrary);

            IAssociationView pAssociationView = pView as IAssociationView;
            pAssociationView.SourceFileName = _part7.FileName;

            IAssociationViewElements assocElems = (IAssociationViewElements)pAssociationView;
            assocElems.CreateCircularCentres = false;
            assocElems.CreateLinearCentres = false;
            assocElems.CreateAxis = false;
            assocElems.CreateCentresMarkers = false;
            assocElems.ProjectAxis = false;
            assocElems.ProjectDesTexts = false;
            assocElems.ProjectSpecRough = false;

            IEmbodimentsManager emb = (IEmbodimentsManager)pAssociationView;
            emb.SetCurrentEmbodiment(_part7.Marking);

            pAssociationView.Angle = 0;
            pAssociationView.X = 0;
            pAssociationView.Y = 0;
            pAssociationView.BendLinesVisible = false;
            pAssociationView.BreakLinesVisible = false;
            pAssociationView.HiddenLinesVisible = false;
            pAssociationView.VisibleLinesStyle = (int)ksCurveStyleEnum.ksCSNormal;
            pAssociationView.Scale = 1;
            pAssociationView.Name = "User view";
            pAssociationView.ProjectionName = "Авторазвертка";
            pAssociationView.Unfold = true;
            pAssociationView.CenterLinesVisible = false;
            pAssociationView.SourceFileName = _part7.FileName;
            pAssociationView.Update();
            pView.Update();

            IViewDesignation pViewDesignation = pView as IViewDesignation;
            pViewDesignation.ShowUnfold = false;
            pViewDesignation.ShowScale = false;

            pView.Update();
            document2D.ksRebuildDocument();
            _application.HideMessage = ksHideMessageEnum.ksShowMessage;

            document2D.ksSaveDocument(savePath);
            if (savePath2 != null)
                document2D.ksSaveDocument(savePath2);

            IKompasDocument kompasDocument = (IKompasDocument)_application.ActiveDocument;
            kompasDocument.Close(DocumentCloseOptions.kdDoNotSaveChanges);
        }

        // Event sink — реализация ksSelectionMngNotify для получения событий выделения
        private class SelectionSink : ksSelectionMngNotify
        {
            private readonly FaceSelectForm _form;
            private readonly ksSelectionMng _selMng;

            public SelectionSink(FaceSelectForm form, ksSelectionMng selMng)
            {
                _form = form;
                _selMng = selMng;
            }

            public bool Select(object obj)
            {
                try
                {
                    ksEntity entity = obj as ksEntity;
                    if (entity != null)
                    {
                        if (entity.IsIt((int)Obj3dType.o3d_face))
                            _form.OnFaceSelected(obj);
                        else
                            _form.OnWrongTypeSelected(entity.type);
                        return true;
                    }
                    // fallback: перебрать менеджер
                    int count = _selMng.GetCount();
                    for (int i = 0; i < count; i++)
                    {
                        if (_selMng.GetObjectType(i) == (int)Obj3dType.o3d_face)
                        {
                            _form.OnFaceSelected(_selMng.GetObjectByIndex(i));
                            return true;
                        }
                    }
                    _form.OnWrongTypeSelected(count > 0 ? _selMng.GetObjectType(0) : -1);
                }
                catch { }
                return true;
            }

            public bool Unselect(object obj)
            {
                _form.OnSelectionCleared();
                return true;
            }

            public bool UnselectAll()
            {
                _form.OnSelectionCleared();
                return true;
            }
        }
    }
}
