namespace SetPrices
{
    partial class SetPriceForm
    {
        /// <summary>
        /// Требуется переменная конструктора.
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
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.label2 = new System.Windows.Forms.Label();
            this.tbTextFile = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.cbNoTranslate = new System.Windows.Forms.CheckBox();
            this.cbOnlyPrices = new System.Windows.Forms.CheckBox();
            this.cbBrand = new System.Windows.Forms.ComboBox();
            this.cbSeason = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(147, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Выберите файл с товарами";
            // 
            // tbTextFile
            // 
            this.tbTextFile.Location = new System.Drawing.Point(184, 13);
            this.tbTextFile.Name = "tbTextFile";
            this.tbTextFile.ReadOnly = true;
            this.tbTextFile.Size = new System.Drawing.Size(262, 20);
            this.tbTextFile.TabIndex = 3;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(452, 13);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(29, 21);
            this.button2.TabIndex = 5;
            this.button2.Text = "...";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(322, 79);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(159, 30);
            this.button3.TabIndex = 6;
            this.button3.Text = "Создать файл";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // cbNoTranslate
            // 
            this.cbNoTranslate.AutoSize = true;
            this.cbNoTranslate.Location = new System.Drawing.Point(184, 87);
            this.cbNoTranslate.Name = "cbNoTranslate";
            this.cbNoTranslate.Size = new System.Drawing.Size(132, 17);
            this.cbNoTranslate.TabIndex = 7;
            this.cbNoTranslate.Text = "не переводить цвета";
            this.cbNoTranslate.UseVisualStyleBackColor = true;
            // 
            // cbOnlyPrices
            // 
            this.cbOnlyPrices.AutoSize = true;
            this.cbOnlyPrices.Location = new System.Drawing.Point(88, 87);
            this.cbOnlyPrices.Name = "cbOnlyPrices";
            this.cbOnlyPrices.Size = new System.Drawing.Size(90, 17);
            this.cbOnlyPrices.TabIndex = 8;
            this.cbOnlyPrices.Text = "только цены";
            this.cbOnlyPrices.UseVisualStyleBackColor = true;
            this.cbOnlyPrices.CheckedChanged += new System.EventHandler(this.cbOnlyPrices_CheckedChanged);
            // 
            // cbBrand
            // 
            this.cbBrand.FormattingEnabled = true;
            this.cbBrand.Items.AddRange(new object[] {
            "ABSOLU",
            "BAGUTTI",
            "BENINO",
            "CANNELLA",
            "CHERIE",
            "Eden-Rose",
            "ELISA-FANTI",
            "ETINCELLE",
            "FLORENCE-MODE",
            "GEORGEDE",
            "GIANI-FORTE",
            "GREGE",
            "HERESIS",
            "JEAN-MARC-PHILIPPE",
            "KOR-KOR",
            "LEO-GUY",
            "LEO&UGO",
            "ME",
            "MITIKA",
            "ORNA-FARHO",
            "ODEMAY",
            "P.CARAT",
            "PASSIONI",
            "PATOUCHKA",
            "ReneDerhy",
            "RINATTI",
            "SAGAI",
            "SPARKLE",
            "TIZIANO-SANTANDREA",
            "MISS-J"});
            this.cbBrand.Location = new System.Drawing.Point(16, 46);
            this.cbBrand.Name = "cbBrand";
            this.cbBrand.Size = new System.Drawing.Size(300, 21);
            this.cbBrand.TabIndex = 9;
            // 
            // cbSeason
            // 
            this.cbSeason.FormattingEnabled = true;
            this.cbSeason.Items.AddRange(new object[] {
            "зима 2019",
            "лето 2019",
            "зима 2020",
            "лето 2020"});
            this.cbSeason.Location = new System.Drawing.Point(322, 46);
            this.cbSeason.Name = "cbSeason";
            this.cbSeason.Size = new System.Drawing.Size(159, 21);
            this.cbSeason.TabIndex = 10;
            // 
            // SetPriceForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(489, 120);
            this.Controls.Add(this.cbSeason);
            this.Controls.Add(this.cbBrand);
            this.Controls.Add(this.cbOnlyPrices);
            this.Controls.Add(this.cbNoTranslate);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.tbTextFile);
            this.Controls.Add(this.label2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SetPriceForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Импорт номенклатур";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbTextFile;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.CheckBox cbNoTranslate;
        private System.Windows.Forms.CheckBox cbOnlyPrices;
        private System.Windows.Forms.ComboBox cbBrand;
        private System.Windows.Forms.ComboBox cbSeason;
    }
}

