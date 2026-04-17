using System;
using System.IO;
using System.Windows.Forms;
using PdfiumViewer;

namespace Multitool
{
    class PdfViewerControl : UserControl
    {
        private PdfViewer pdfViewer;

        public PdfViewerControl()
        {
            this.pdfViewer = new PdfViewer();
            this.pdfViewer.Dock = DockStyle.Fill;
            this.Controls.Add(this.pdfViewer);
            this.Name = "PdfViewerControl";
            this.Size = new System.Drawing.Size(600, 400);
        }

        public void LoadPdf(string filePath)
        {
            if (!File.Exists(filePath))
            {
                MessageBox.Show("Файл не найден: " + filePath);
                return;
            }

            pdfViewer.Document?.Dispose();
            pdfViewer.Document = null;

            // Читаем в память — оригинальный файл не блокируется и остаётся
            // доступным для перезаписи или удаления
            byte[] data = File.ReadAllBytes(filePath);

            // В режиме Library (COM DLL) Handle контрола может не быть создан
            // в момент вызова — Renderer.Load() → Invalidate() уходит в никуда
            if (!pdfViewer.IsHandleCreated)
                pdfViewer.CreateControl();

            pdfViewer.Document = PdfDocument.Load(new MemoryStream(data));
            pdfViewer.Refresh();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                pdfViewer?.Document?.Dispose();

            base.Dispose(disposing);
        }
    }
}
