using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using netDxf;
using netDxf.Entities;
using Point = System.Drawing.Point;

namespace Multitool
{
    /// <summary>
    /// Просмотрщик DXF-файлов на GDI+. Не требует установки eDrawings.
    /// Зум — колесо мыши. Пан — зажатая левая кнопка.
    /// </summary>
    public class DxfViewerControl : Control
    {
        private DxfDocument _doc;
        private RectangleF _bounds = RectangleF.Empty;

        private float _scale = 1f;
        private float _panX, _panY;

        private bool _panning;
        private Point _lastMouse;

        public DxfViewerControl()
        {
            DoubleBuffered = true;
            BackColor = Color.White;
        }

        public void LoadDxf(string path)
        {
            try   { _doc = DxfDocument.Load(path); }
            catch { _doc = null; }
            ComputeBounds();
            FitToView();
            Invalidate();
        }

        public void Clear()
        {
            _doc = null;
            _bounds = RectangleF.Empty;
            Invalidate();
        }

        // ────────────────────── Bounds ──────────────────────────────────────────

        private void ComputeBounds()
        {
            if (_doc == null) { _bounds = RectangleF.Empty; return; }

            float minX = float.MaxValue, minY = float.MaxValue;
            float maxX = float.MinValue, maxY = float.MinValue;

            void Expand(double x, double y)
            {
                float fx = (float)x, fy = (float)y;
                if (fx < minX) minX = fx; if (fx > maxX) maxX = fx;
                if (fy < minY) minY = fy; if (fy > maxY) maxY = fy;
            }

            foreach (var e in _doc.Lines)
            {
                Expand(e.StartPoint.X, e.StartPoint.Y);
                Expand(e.EndPoint.X, e.EndPoint.Y);
            }
            foreach (var e in _doc.Arcs)
            {
                Expand(e.Center.X - e.Radius, e.Center.Y - e.Radius);
                Expand(e.Center.X + e.Radius, e.Center.Y + e.Radius);
            }
            foreach (var e in _doc.Circles)
            {
                Expand(e.Center.X - e.Radius, e.Center.Y - e.Radius);
                Expand(e.Center.X + e.Radius, e.Center.Y + e.Radius);
            }
            foreach (var e in _doc.LwPolylines)
                foreach (var v in e.Vertexes)
                    Expand(v.Position.X, v.Position.Y);
            foreach (var e in _doc.Ellipses)
            {
                double r = Math.Max(e.MajorAxis, e.MinorAxis);
                Expand(e.Center.X - r, e.Center.Y - r);
                Expand(e.Center.X + r, e.Center.Y + r);
            }

            _bounds = (minX < maxX && minY < maxY)
                ? RectangleF.FromLTRB(minX, minY, maxX, maxY)
                : RectangleF.Empty;
        }

        private void FitToView()
        {
            if (_bounds.IsEmpty || Width == 0 || Height == 0) return;
            const float margin = 20;
            float sx = (Width  - margin * 2) / _bounds.Width;
            float sy = (Height - margin * 2) / _bounds.Height;
            _scale = Math.Min(sx, sy);
            _panX  = margin - _bounds.Left   * _scale + (Width  - margin * 2 - _bounds.Width  * _scale) / 2f;
            _panY  = margin + _bounds.Bottom * _scale + (Height - margin * 2 - _bounds.Height * _scale) / 2f;
        }

        // ────────────────────── Рендеринг ───────────────────────────────────────

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            if (_doc == null)
            {
                using (var f = new Font("Segoe UI", 10))
                    g.DrawString("DXF не загружен", f, Brushes.Gray, 10, 10);
                return;
            }

            g.TranslateTransform(_panX, _panY);
            g.ScaleTransform(_scale, -_scale);

            using (var pen = new Pen(Color.FromArgb(0, 120, 215), 1f / _scale))
            {
                foreach (var e2 in _doc.Lines)
                    g.DrawLine(pen,
                        (float)e2.StartPoint.X, (float)e2.StartPoint.Y,
                        (float)e2.EndPoint.X,   (float)e2.EndPoint.Y);

                foreach (var e2 in _doc.Circles)
                {
                    float r = (float)e2.Radius;
                    g.DrawEllipse(pen,
                        (float)e2.Center.X - r, (float)e2.Center.Y - r,
                        r * 2, r * 2);
                }

                foreach (var e2 in _doc.Arcs)
                    DrawArc(g, pen, e2);

                foreach (var e2 in _doc.LwPolylines)
                    DrawLwPolyline(g, pen, e2);

                foreach (var e2 in _doc.Ellipses)
                    DrawEllipse(g, pen, e2);
            }
        }

        private static void DrawArc(Graphics g, Pen pen, Arc a)
        {
            float r = (float)a.Radius;
            float start = (float)a.StartAngle;
            float end   = (float)a.EndAngle;
            float sweep = end - start;
            if (sweep <= 0) sweep += 360f;
            g.DrawArc(pen, (float)a.Center.X - r, (float)a.Center.Y - r, r * 2, r * 2, start, sweep);
        }

        private static void DrawLwPolyline(Graphics g, Pen pen, LwPolyline p)
        {
            var verts = p.Vertexes;
            if (verts.Count < 2) return;
            for (int k = 0; k < verts.Count - 1; k++)
                g.DrawLine(pen,
                    (float)verts[k].Position.X,     (float)verts[k].Position.Y,
                    (float)verts[k + 1].Position.X, (float)verts[k + 1].Position.Y);
            if (p.IsClosed)
                g.DrawLine(pen,
                    (float)verts[verts.Count - 1].Position.X,
                    (float)verts[verts.Count - 1].Position.Y,
                    (float)verts[0].Position.X, (float)verts[0].Position.Y);
        }

        private static void DrawEllipse(Graphics g, Pen pen, Ellipse el)
        {
            double maj = el.MajorAxis;
            double min = el.MinorAxis;
            double rot = el.Rotation * Math.PI / 180.0;
            double startDeg = el.IsFullEllipse ? 0 : el.StartAngle;
            double endDeg   = el.IsFullEllipse ? 360 : el.EndAngle;
            double totalDeg = endDeg - startDeg;
            if (totalDeg <= 0) totalDeg += 360;

            const int steps = 72;
            PointF Pt(double deg)
            {
                double t  = deg * Math.PI / 180.0;
                double ex = maj * Math.Cos(t);
                double ey = min * Math.Sin(t);
                return new PointF(
                    (float)(ex * Math.Cos(rot) - ey * Math.Sin(rot) + el.Center.X),
                    (float)(ex * Math.Sin(rot) + ey * Math.Cos(rot) + el.Center.Y));
            }

            PointF prev = Pt(startDeg);
            for (int k = 1; k <= steps; k++)
            {
                PointF cur = Pt(startDeg + totalDeg * k / steps);
                g.DrawLine(pen, prev, cur);
                prev = cur;
            }
        }

        // ────────────────────── Мышь ────────────────────────────────────────────

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            float factor = e.Delta > 0 ? 1.15f : 1f / 1.15f;
            _panX = e.X + (_panX - e.X) * factor;
            _panY = e.Y + (_panY - e.Y) * factor;
            _scale *= factor;
            Invalidate();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button == MouseButtons.Left) { _panning = true; _lastMouse = e.Location; Cursor = Cursors.Hand; }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_panning) { _panX += e.X - _lastMouse.X; _panY += e.Y - _lastMouse.Y; _lastMouse = e.Location; Invalidate(); }
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _panning = false;
            Cursor = Cursors.Default;
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (_doc != null && !_bounds.IsEmpty) FitToView();
            Invalidate();
        }
    }
}
