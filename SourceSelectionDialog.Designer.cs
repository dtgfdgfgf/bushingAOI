// 修正 RadioButton 容器配置

namespace peilin
{
    partial class SourceSelectionDialog
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
            // 移除 rbSuggested（將其改到 Form 層級）
            // this.groupBox1.Controls.Add(this.rbSuggested); // ← 移除這行
            this.groupBox1.Location = new System.Drawing.Point(12, 32);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(460, 180);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "";
            // 
            // panelSuggested
            // 
            this.panelSuggested.AutoScroll = true;
            this.panelSuggested.Location = new System.Drawing.Point(20, 45);
            this.panelSuggested.Name = "panelSuggested";
            this.panelSuggested.Size = new System.Drawing.Size(420, 120);
            this.panelSuggested.TabIndex = 2;
            // 
            // lblSuggestion
            // 
            this.lblSuggestion.AutoSize = true;
            this.lblSuggestion.Location = new System.Drawing.Point(20, 25);
            this.lblSuggestion.Name = "lblSuggestion";
            this.lblSuggestion.Size = new System.Drawing.Size(137, 12);
            this.lblSuggestion.TabIndex = 1;
            this.lblSuggestion.Text = "相同PTFE的現有料號:";
            // 
            // rbSuggested
            // 
            this.rbSuggested.AutoSize = true;
            this.rbSuggested.Location = new System.Drawing.Point(22, 12);
            this.rbSuggested.Name = "rbSuggested";
            this.rbSuggested.Size = new System.Drawing.Size(95, 16);
            this.rbSuggested.TabIndex = 0;
            this.rbSuggested.TabStop = true;
            this.rbSuggested.Text = "使用建議料號";
            this.rbSuggested.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cmbAllTypes);
            // 移除 rbManual（將其改到 Form 層級）
            // this.groupBox2.Controls.Add(this.rbManual); // ← 移除這行
            this.groupBox2.Location = new System.Drawing.Point(12, 238);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(460, 60);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "";
            // 
            // cmbAllTypes
            // 
            this.cmbAllTypes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAllTypes.FormattingEnabled = true;
            this.cmbAllTypes.Location = new System.Drawing.Point(20, 25);
            this.cmbAllTypes.Name = "cmbAllTypes";
            this.cmbAllTypes.Size = new System.Drawing.Size(420, 20);
            this.cmbAllTypes.TabIndex = 1;
            // 
            // rbManual
            // 
            this.rbManual.AutoSize = true;
            this.rbManual.Location = new System.Drawing.Point(22, 218);
            this.rbManual.Name = "rbManual";
            this.rbManual.Size = new System.Drawing.Size(119, 16);
            this.rbManual.TabIndex = 2;
            this.rbManual.TabStop = true;
            this.rbManual.Text = "手動選擇來源料號";
            this.rbManual.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(316, 310);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 30);
            this.btnOK.TabIndex = 4;
            this.btnOK.Text = "確定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(397, 310);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 30);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // SourceSelectionDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 352);
            // 將 RadioButton 加入 Form（關鍵修正）
            this.Controls.Add(this.rbSuggested);
            this.Controls.Add(this.rbManual);
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
            this.ResumeLayout(false);
            this.PerformLayout();

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