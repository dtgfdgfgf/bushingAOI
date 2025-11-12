using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace peilin
{
    partial class ParameterConfigForm
    {
        private System.ComponentModel.IContainer components = null;
        private TabControl tabMain;
        private TabPage tabCamera;
        private TabPage tabPosition;
        private TabPage tabDetection;
        private TabPage tabTiming;
        private TabPage tabTesting;

        // 進度顯示控件
        private Label lblProgress;
        private ProgressBar progressBarOverall;
        private Label lblProgressSummary;

        // 相機參數Tab的控件
        private DataGridView dgvCameraUnmodified;
        private DataGridView dgvCameraModified;
        private DataGridView dgvCameraFixed;
        private GroupBox grpCameraUnmodified;
        private GroupBox grpCameraModified;
        private GroupBox grpCameraFixed;
        private Button btnCameraMoveToModified;
        private Button btnCameraMoveToFixed;
        private Button btnCameraMoveToUnmodified;
        private Button btnCameraSelectAll;
        private Button btnCameraSelectNone;
        private Button btnCameraSelectInvert;

        // 位置參數Tab的控件
        private DataGridView dgvPositionUnmodified;
        private DataGridView dgvPositionModified;
        private DataGridView dgvPositionFixed;
        private GroupBox grpPositionUnmodified;
        private GroupBox grpPositionModified;
        private GroupBox grpPositionFixed;
        private Button btnPositionMoveToModified;
        private Button btnPositionMoveToFixed;
        private Button btnPositionMoveToUnmodified;
        private Button btnPositionSelectAll;
        private Button btnPositionSelectNone;
        private Button btnPositionSelectInvert;
        private Button btnOpenPositionCalibration;
        private Label lblPositionHint;

        // 檢測參數Tab的控件
        private DataGridView dgvDetectionUnmodified;
        private DataGridView dgvDetectionModified;
        private DataGridView dgvDetectionFixed;
        private GroupBox grpDetectionUnmodified;
        private GroupBox grpDetectionModified;
        private GroupBox grpDetectionFixed;
        private Button btnDetectionMoveToModified;
        private Button btnDetectionMoveToFixed;
        private Button btnDetectionMoveToUnmodified;
        private Button btnDetectionSelectAll;
        private Button btnDetectionSelectNone;
        private Button btnDetectionSelectInvert;

        // 時間參數Tab的控件
        private DataGridView dgvTimingUnmodified;
        private DataGridView dgvTimingModified;
        private DataGridView dgvTimingFixed;
        private GroupBox grpTimingUnmodified;
        private GroupBox grpTimingModified;
        private GroupBox grpTimingFixed;
        private Button btnTimingMoveToModified;
        private Button btnTimingMoveToFixed;
        private Button btnTimingMoveToUnmodified;
        private Button btnTimingSelectAll;
        private Button btnTimingSelectNone;
        private Button btnTimingSelectInvert;

        // 底部按鈕
        private Button btnCancel;
        private Button btnSaveCurrentTab;
        private Button btnSaveAllAndComplete;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tabMain = new System.Windows.Forms.TabControl();
            this.tabCamera = new System.Windows.Forms.TabPage();
            this.tabPosition = new System.Windows.Forms.TabPage();
            this.tabDetection = new System.Windows.Forms.TabPage();
            this.tabTiming = new System.Windows.Forms.TabPage();
            this.tabTesting = new System.Windows.Forms.TabPage();
            this.lblProgress = new System.Windows.Forms.Label();
            this.progressBarOverall = new System.Windows.Forms.ProgressBar();
            this.lblProgressSummary = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSaveCurrentTab = new System.Windows.Forms.Button();
            this.btnSaveAllAndComplete = new System.Windows.Forms.Button();
            this.tabMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabMain
            // 
            this.tabMain.Controls.Add(this.tabCamera);
            this.tabMain.Controls.Add(this.tabPosition);
            this.tabMain.Controls.Add(this.tabDetection);
            this.tabMain.Controls.Add(this.tabTiming);
            //this.tabMain.Controls.Add(this.tabTesting);
            this.tabMain.Location = new System.Drawing.Point(9, 40);
            this.tabMain.Margin = new System.Windows.Forms.Padding(2);
            this.tabMain.Name = "tabMain";
            this.tabMain.SelectedIndex = 0;
            this.tabMain.Size = new System.Drawing.Size(1573, 901);
            this.tabMain.TabIndex = 3;
            this.tabMain.SelectedIndexChanged += new System.EventHandler(this.tabMain_SelectedIndexChanged);
            // 
            // tabCamera
            // 
            this.tabCamera.Location = new System.Drawing.Point(4, 22);
            this.tabCamera.Margin = new System.Windows.Forms.Padding(2);
            this.tabCamera.Name = "tabCamera";
            this.tabCamera.Size = new System.Drawing.Size(1565, 875);
            this.tabCamera.TabIndex = 0;
            this.tabCamera.Text = "相機參數";
            this.tabCamera.UseVisualStyleBackColor = true;
            // 
            // tabPosition
            // 
            this.tabPosition.Location = new System.Drawing.Point(4, 22);
            this.tabPosition.Margin = new System.Windows.Forms.Padding(2);
            this.tabPosition.Name = "tabPosition";
            this.tabPosition.Size = new System.Drawing.Size(1565, 833);
            this.tabPosition.TabIndex = 1;
            this.tabPosition.Text = "位置參數";
            this.tabPosition.UseVisualStyleBackColor = true;
            // 
            // tabDetection
            // 
            this.tabDetection.Location = new System.Drawing.Point(4, 22);
            this.tabDetection.Margin = new System.Windows.Forms.Padding(2);
            this.tabDetection.Name = "tabDetection";
            this.tabDetection.Size = new System.Drawing.Size(1565, 833);
            this.tabDetection.TabIndex = 2;
            this.tabDetection.Text = "檢測參數";
            this.tabDetection.UseVisualStyleBackColor = true;
            // 
            // tabTiming
            // 
            this.tabTiming.Location = new System.Drawing.Point(4, 22);
            this.tabTiming.Margin = new System.Windows.Forms.Padding(2);
            this.tabTiming.Name = "tabTiming";
            this.tabTiming.Size = new System.Drawing.Size(1565, 833);
            this.tabTiming.TabIndex = 3;
            this.tabTiming.Text = "時間參數";
            this.tabTiming.UseVisualStyleBackColor = true;
            // 
            // tabTesting
            // 
            this.tabTesting.Location = new System.Drawing.Point(4, 22);
            this.tabTesting.Margin = new System.Windows.Forms.Padding(2);
            this.tabTesting.Name = "tabTesting";
            this.tabTesting.Size = new System.Drawing.Size(1565, 833);
            this.tabTesting.TabIndex = 4;
            this.tabTesting.Text = "測試驗證";
            this.tabTesting.UseVisualStyleBackColor = true;
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.lblProgress.Location = new System.Drawing.Point(9, 12);
            this.lblProgress.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(122, 21);
            this.lblProgress.TabIndex = 0;
            this.lblProgress.Text = "料號: [載入中...]";
            // 
            // progressBarOverall
            // 
            this.progressBarOverall.Location = new System.Drawing.Point(165, 14);
            this.progressBarOverall.Margin = new System.Windows.Forms.Padding(2);
            this.progressBarOverall.Name = "progressBarOverall";
            this.progressBarOverall.Size = new System.Drawing.Size(225, 16);
            this.progressBarOverall.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBarOverall.TabIndex = 1;
            // 
            // lblProgressSummary
            // 
            this.lblProgressSummary.AutoSize = true;
            this.lblProgressSummary.Location = new System.Drawing.Point(398, 16);
            this.lblProgressSummary.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProgressSummary.Name = "lblProgressSummary";
            this.lblProgressSummary.Size = new System.Drawing.Size(312, 12);
            this.lblProgressSummary.TabIndex = 2;
            this.lblProgressSummary.Text = "❌相機(0/0) ❌位置(0/0) ❌檢測(0/0) ❌時間(0/0) ❌測試(0/0)";
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(1293, 961);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(2);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 28);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "取消";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSaveCurrentTab
            // 
            this.btnSaveCurrentTab.BackColor = System.Drawing.Color.LightGreen;
            this.btnSaveCurrentTab.Location = new System.Drawing.Point(1375, 961);
            this.btnSaveCurrentTab.Margin = new System.Windows.Forms.Padding(2);
            this.btnSaveCurrentTab.Name = "btnSaveCurrentTab";
            this.btnSaveCurrentTab.Size = new System.Drawing.Size(90, 28);
            this.btnSaveCurrentTab.TabIndex = 5;
            this.btnSaveCurrentTab.Text = "儲存當前頁面";
            this.btnSaveCurrentTab.UseVisualStyleBackColor = false;
            this.btnSaveCurrentTab.Click += new System.EventHandler(this.btnSaveCurrentTab_Click);
            // 
            // btnSaveAllAndComplete
            // 
            this.btnSaveAllAndComplete.BackColor = System.Drawing.Color.LightCoral;
            this.btnSaveAllAndComplete.Location = new System.Drawing.Point(1473, 961);
            this.btnSaveAllAndComplete.Margin = new System.Windows.Forms.Padding(2);
            this.btnSaveAllAndComplete.Name = "btnSaveAllAndComplete";
            this.btnSaveAllAndComplete.Size = new System.Drawing.Size(105, 28);
            this.btnSaveAllAndComplete.TabIndex = 6;
            this.btnSaveAllAndComplete.Text = "儲存全部並完成";
            this.btnSaveAllAndComplete.UseVisualStyleBackColor = false;
            this.btnSaveAllAndComplete.Click += new System.EventHandler(this.btnSaveAllAndComplete_Click);
            // 
            // ParameterConfigForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1593, 1009);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.progressBarOverall);
            this.Controls.Add(this.lblProgressSummary);
            this.Controls.Add(this.tabMain);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSaveCurrentTab);
            this.Controls.Add(this.btnSaveAllAndComplete);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MinimizeBox = false;
            this.Name = "ParameterConfigForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "新料號參數設定";
            this.tabMain.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}