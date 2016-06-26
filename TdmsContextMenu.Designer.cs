namespace Auto
{
    partial class Property
    {
        private System.ComponentModel.IContainer components = null;
             
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Property));
            this.chboxUpdateScheduleTable = new System.Windows.Forms.CheckBox();
            this.chboxUpdateAttr = new System.Windows.Forms.CheckBox();
            this.chboxUpdateXref = new System.Windows.Forms.CheckBox();
            this.chboxSearchChangeXref = new System.Windows.Forms.CheckBox();
            this.cmdOK = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.gbXref = new System.Windows.Forms.GroupBox();
            this.gbTrace = new System.Windows.Forms.GroupBox();
            this.tbTracePath = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnSetPath = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.gbXref.SuspendLayout();
            this.gbTrace.SuspendLayout();
            this.SuspendLayout();
            // 
            // chboxUpdateScheduleTable
            // 
            this.chboxUpdateScheduleTable.AutoSize = true;
            this.chboxUpdateScheduleTable.Location = new System.Drawing.Point(7, 76);
            this.chboxUpdateScheduleTable.Name = "chboxUpdateScheduleTable";
            this.chboxUpdateScheduleTable.Size = new System.Drawing.Size(203, 17);
            this.chboxUpdateScheduleTable.TabIndex = 8;
            this.chboxUpdateScheduleTable.Text = "Обновление таблиц спецификаций";
            this.chboxUpdateScheduleTable.UseVisualStyleBackColor = true;
            // 
            // chboxUpdateAttr
            // 
            this.chboxUpdateAttr.AutoSize = true;
            this.chboxUpdateAttr.Location = new System.Drawing.Point(7, 53);
            this.chboxUpdateAttr.Name = "chboxUpdateAttr";
            this.chboxUpdateAttr.Size = new System.Drawing.Size(190, 17);
            this.chboxUpdateAttr.TabIndex = 7;
            this.chboxUpdateAttr.Text = "Обновление атрибутов штампов";
            this.chboxUpdateAttr.UseVisualStyleBackColor = true;
            // 
            // chboxUpdateXref
            // 
            this.chboxUpdateXref.AutoSize = true;
            this.chboxUpdateXref.Checked = true;
            this.chboxUpdateXref.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chboxUpdateXref.Location = new System.Drawing.Point(7, 30);
            this.chboxUpdateXref.Name = "chboxUpdateXref";
            this.chboxUpdateXref.Size = new System.Drawing.Size(175, 17);
            this.chboxUpdateXref.TabIndex = 6;
            this.chboxUpdateXref.Text = "Обновление внешних ссылок";
            this.chboxUpdateXref.UseVisualStyleBackColor = true;
            // 
            // chboxSearchChangeXref
            // 
            this.chboxSearchChangeXref.AutoSize = true;
            this.chboxSearchChangeXref.Location = new System.Drawing.Point(7, 99);
            this.chboxSearchChangeXref.Name = "chboxSearchChangeXref";
            this.chboxSearchChangeXref.Size = new System.Drawing.Size(222, 17);
            this.chboxSearchChangeXref.TabIndex = 9;
            this.chboxSearchChangeXref.Text = "Проверка изменений внешних ссылок";
            this.chboxSearchChangeXref.UseVisualStyleBackColor = true;
            // 
            // cmdOK
            // 
            this.cmdOK.Location = new System.Drawing.Point(179, 217);
            this.cmdOK.Name = "cmdOK";
            this.cmdOK.Size = new System.Drawing.Size(196, 30);
            this.cmdOK.TabIndex = 10;
            this.cmdOK.Text = "Применить  изменения и закрыть";
            this.cmdOK.UseVisualStyleBackColor = true;
            this.cmdOK.Click += new System.EventHandler(this.button1_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(9, 1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(222, 50);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 11;
            this.pictureBox1.TabStop = false;
            // 
            // gbXref
            // 
            this.gbXref.Controls.Add(this.chboxUpdateAttr);
            this.gbXref.Controls.Add(this.chboxUpdateXref);
            this.gbXref.Controls.Add(this.chboxUpdateScheduleTable);
            this.gbXref.Controls.Add(this.chboxSearchChangeXref);
            this.gbXref.Location = new System.Drawing.Point(9, 57);
            this.gbXref.Name = "gbXref";
            this.gbXref.Size = new System.Drawing.Size(260, 154);
            this.gbXref.TabIndex = 12;
            this.gbXref.TabStop = false;
            this.gbXref.Text = "Настройки внешних ссылок";
            // 
            // gbTrace
            // 
            this.gbTrace.Controls.Add(this.tbTracePath);
            this.gbTrace.Controls.Add(this.label2);
            this.gbTrace.Controls.Add(this.btnSetPath);
            this.gbTrace.Location = new System.Drawing.Point(279, 57);
            this.gbTrace.Name = "gbTrace";
            this.gbTrace.Size = new System.Drawing.Size(260, 154);
            this.gbTrace.TabIndex = 13;
            this.gbTrace.TabStop = false;
            this.gbTrace.Text = "Настройки логирования";
            // 
            // tbTracePath
            // 
            this.tbTracePath.Location = new System.Drawing.Point(6, 42);
            this.tbTracePath.Name = "tbTracePath";
            this.tbTracePath.Size = new System.Drawing.Size(248, 20);
            this.tbTracePath.TabIndex = 17;
            this.tbTracePath.Text = TraceSource.LogFileFullPath;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(155, 13);
            this.label2.TabIndex = 16;
            this.label2.Text = "Текущий путь к файлу логов:";
            // 
            // btnSetPath
            // 
            this.btnSetPath.Location = new System.Drawing.Point(171, 69);
            this.btnSetPath.Name = "btnSetPath";
            this.btnSetPath.Size = new System.Drawing.Size(83, 24);
            this.btnSetPath.TabIndex = 14;
            this.btnSetPath.Text = "Задать путь";
            this.btnSetPath.UseVisualStyleBackColor = true;
            this.btnSetPath.Click += new System.EventHandler(this.btnSetPath_Click);
            // 
            // Property
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(549, 254);
            this.Controls.Add(this.gbTrace);
            this.Controls.Add(this.gbXref);
            this.Controls.Add(this.cmdOK);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Property";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройки";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.gbXref.ResumeLayout(false);
            this.gbXref.PerformLayout();
            this.gbTrace.ResumeLayout(false);
            this.gbTrace.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckBox chboxUpdateScheduleTable;
        private System.Windows.Forms.CheckBox chboxUpdateAttr;
        private System.Windows.Forms.CheckBox chboxUpdateXref;
        private System.Windows.Forms.CheckBox chboxSearchChangeXref;
        private System.Windows.Forms.Button cmdOK;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.GroupBox gbXref;
        private System.Windows.Forms.GroupBox gbTrace;
        private System.Windows.Forms.TextBox tbTracePath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnSetPath;
    }
}