namespace peilin
{
    partial class SourceSelectionDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panelSuggested = new System.Windows.Forms.Panel();
            this.lblSuggestion = new System.Windows.Forms.Label();
            this.rbSuggested = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cmbAllTypes = new System.Windows.Forms.ComboBox();
            this.rbManual = new System.Windows.Forms.RadioButton();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panelSuggested);
            this.groupBox1.Controls.Add(this.lblSuggestion);
            this.groupBox1.Controls.Add(this.rbSuggested);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(460, 180);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "自動建議";
            // 
            // panelSuggested
            // 
            this.panelSuggested.AutoScroll = true;
            this.panelSuggested.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelSuggested.Location = new System.Drawing.Point(20, 65);
            this.panelSuggested.Name = "panelSuggested";
            this.panelSuggested.Size = new System.Drawing.Size(420, 100);
            this.panelSuggested.TabIndex = 2;
            // 
            // lblSuggestion
            // 
            this.lblSuggestion.AutoSize = true;
            this.lblSuggestion.Location = new System.Drawing.Point(17, 45);
            this.lblSuggestion.Name = "lblSuggestion";
            this.lblSuggestion.Size = new System.Drawing.Size(105, 13);
            this.lblSuggestion.TabIndex = 1;
            this.lblSuggestion.Text = "相同PTFE的料號:";
            // 
            // rbSuggested
            // 
            this.rbSuggested.AutoSize = true;
            this.rbSuggested.Checked = true;
            this.rbSuggested.Location = new System.Drawing.Point(6, 19);
            this.rbSuggested.Name = "rbSuggested";
            this.rbSuggested.Size = new System.Drawing.Size(85, 17);
            this.rbSuggested.TabIndex = 0;
            this.rbSuggested.TabStop = true;
            this.rbSuggested.Text = "使用建議料號";
            this.rbSuggested.UseVisualStyleBackColor = true;
            this.rbSuggested.CheckedChanged += new System.EventHandler(this.rbSuggested_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cmbAllTypes);
            this.groupBox2.Controls.Add(this.rbManual);
            this.groupBox2.Location = new System.Drawing.Point(12, 210);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(460, 80);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "手動選擇";
            // 
            // cmbAllTypes
            // 
            this.cmbAllTypes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAllTypes.Enabled = false;
            this.cmbAllTypes.FormattingEnabled = true;
            this.cmbAllTypes.Location = new System.Drawing.Point(20, 45);
            this.cmbAllTypes.Name = "cmbAllTypes";
            this.cmbAllTypes.Size = new System.Drawing.Size(420, 21);
            this.cmbAllTypes.TabIndex = 1;
            // 
            // rbManual
            // 
            this.rbManual.AutoSize = true;
            this.rbManual.Location = new System.Drawing.Point(6, 19);
            this.rbManual.Name = "rbManual";
            this.rbManual.Size = new System.Drawing.Size(109, 17);
            this.rbManual.TabIndex = 0;
            this.rbManual.Text = "手動選擇來源料號";
            this.rbManual.UseVisualStyleBackColor = true;
            this.rbManual.CheckedChanged += new System.EventHandler(this.rbManual_CheckedChanged);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(316, 310);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 30);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "確定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(397, 310);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 30);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // SourceSelectionDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 361);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SourceSelectionDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "選擇來源料號";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panelSuggested;
        private System.Windows.Forms.Label lblSuggestion;
        private System.Windows.Forms.RadioButton rbSuggested;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox cmbAllTypes;
        private System.Windows.Forms.RadioButton rbManual;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
    }
}
