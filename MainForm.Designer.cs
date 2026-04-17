
namespace Multitool
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.CreateDXF = new System.Windows.Forms.Button();
            this.CreatePDF = new System.Windows.Forms.Button();
            this.CreateExcel = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.newDXF = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.Settings = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.statusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.tableLayoutPanel1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            //
            // CreateDXF
            //
            this.CreateDXF.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CreateDXF.Location = new System.Drawing.Point(3, 3);
            this.CreateDXF.Name = "CreateDXF";
            this.CreateDXF.Size = new System.Drawing.Size(144, 41);
            this.CreateDXF.TabIndex = 0;
            this.CreateDXF.Text = "Создать DXF";
            this.CreateDXF.UseVisualStyleBackColor = true;
            this.CreateDXF.Visible = false;
            this.CreateDXF.Click += new System.EventHandler(this.СreateDXF_Click);
            //
            // CreatePDF
            //
            this.CreatePDF.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CreatePDF.Location = new System.Drawing.Point(3, 50);
            this.CreatePDF.Name = "CreatePDF";
            this.CreatePDF.Size = new System.Drawing.Size(144, 41);
            this.CreatePDF.TabIndex = 1;
            this.CreatePDF.Text = "Создать PDF";
            this.CreatePDF.UseVisualStyleBackColor = true;
            this.CreatePDF.Click += new System.EventHandler(this.СreatePDF_Click);
            //
            // CreateExcel
            //
            this.CreateExcel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.CreateExcel.Location = new System.Drawing.Point(3, 124);
            this.CreateExcel.Name = "CreateExcel";
            this.CreateExcel.Size = new System.Drawing.Size(144, 41);
            this.CreateExcel.TabIndex = 2;
            this.CreateExcel.Text = "Создать МК Excel";
            this.CreateExcel.UseVisualStyleBackColor = true;
            this.CreateExcel.Click += new System.EventHandler(this.СreateExcel_Click);
            //
            // button4
            //
            this.button4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.button4.Location = new System.Drawing.Point(3, 171);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(144, 41);
            this.button4.TabIndex = 3;
            this.button4.Text = "Исправить модель детали";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            //
            // newDXF
            //
            this.newDXF.Dock = System.Windows.Forms.DockStyle.Fill;
            this.newDXF.Location = new System.Drawing.Point(3, 218);
            this.newDXF.Name = "newDXF";
            this.newDXF.Size = new System.Drawing.Size(144, 41);
            this.newDXF.TabIndex = 4;
            this.newDXF.Text = "Новый DXF";
            this.newDXF.UseVisualStyleBackColor = true;
            this.newDXF.Click += new System.EventHandler(this.newDXF_Click);
            //
            // comboBox1
            //
            this.comboBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Сварочный",
            "Метизный"});
            this.comboBox1.Location = new System.Drawing.Point(3, 97);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(144, 21);
            this.comboBox1.TabIndex = 5;
            //
            // tableLayoutPanel1
            //
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.CreateDXF, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.CreatePDF, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.comboBox1, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.CreateExcel, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.button4, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.newDXF, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.Settings, 0, 8);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 9;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 27F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 12.5F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(826, 398);
            this.tableLayoutPanel1.TabIndex = 6;
            //
            // Settings
            //
            this.Settings.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Settings.Location = new System.Drawing.Point(3, 359);
            this.Settings.Name = "Settings";
            this.Settings.Size = new System.Drawing.Size(144, 41);
            this.Settings.TabIndex = 6;
            this.Settings.Text = "Настройки";
            this.Settings.UseVisualStyleBackColor = true;
            this.Settings.Click += new System.EventHandler(this.Settings_Click);
            //
            // statusLabel
            //
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(0, 17);
            this.statusLabel.Spring = true;
            this.statusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // statusStrip1
            //
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 398);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(826, 22);
            this.statusStrip1.SizingGrip = false;
            this.statusStrip1.TabIndex = 7;
            //
            // MainForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(826, 420);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.statusStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.Text = "Multitool";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button CreateDXF;
        private System.Windows.Forms.Button CreatePDF;
        private System.Windows.Forms.Button CreateExcel;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button newDXF;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button Settings;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel statusLabel;
    }
}
