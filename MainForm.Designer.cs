
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
            this.button5 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // CreateDXF
            // 
            this.CreateDXF.Location = new System.Drawing.Point(13, 13);
            this.CreateDXF.Name = "CreateDXF";
            this.CreateDXF.Size = new System.Drawing.Size(105, 23);
            this.CreateDXF.TabIndex = 0;
            this.CreateDXF.Text = "Создать DXF";
            this.CreateDXF.UseVisualStyleBackColor = true;
            this.CreateDXF.Click += new System.EventHandler(this.СreateDXF_Click);
            // 
            // CreatePDF
            // 
            this.CreatePDF.Location = new System.Drawing.Point(13, 43);
            this.CreatePDF.Name = "CreatePDF";
            this.CreatePDF.Size = new System.Drawing.Size(105, 23);
            this.CreatePDF.TabIndex = 1;
            this.CreatePDF.Text = "Создать PDF";
            this.CreatePDF.UseVisualStyleBackColor = true;
            this.CreatePDF.Click += new System.EventHandler(this.СreatePDF_Click);
            // 
            // CreateExcel
            // 
            this.CreateExcel.Location = new System.Drawing.Point(13, 72);
            this.CreateExcel.Name = "CreateExcel";
            this.CreateExcel.Size = new System.Drawing.Size(105, 23);
            this.CreateExcel.TabIndex = 2;
            this.CreateExcel.Text = "Создать МК Excel";
            this.CreateExcel.UseVisualStyleBackColor = true;
            this.CreateExcel.Click += new System.EventHandler(this.СreateExcel_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(13, 101);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(155, 23);
            this.button4.TabIndex = 3;
            this.button4.Text = "Исправить модель детали";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(13, 130);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 4;
            this.button5.Text = "button5";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Visible = false;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Сварочный",
            "Метизный"});
            this.comboBox1.Location = new System.Drawing.Point(125, 72);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 5;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(259, 164);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.CreateExcel);
            this.Controls.Add(this.CreatePDF);
            this.Controls.Add(this.CreateDXF);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.Text = "Multitool";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button CreateDXF;
        private System.Windows.Forms.Button CreatePDF;
        private System.Windows.Forms.Button CreateExcel;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.ComboBox comboBox1;
    }
}

