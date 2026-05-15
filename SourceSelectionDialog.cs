// 簡化 RadioButton 互斥邏輯

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LinqToDB;

namespace peilin
{
    public partial class SourceSelectionDialog : Form
    {
        public string SelectedSourceType { get; private set; }
        private string targetType;

        // 用於追蹤建議區域選中的料號
        private string selectedSuggestedType = null;

        public SourceSelectionDialog(string targetType)
        {
            InitializeComponent();
            this.targetType = targetType;

            // 訂閱事件
            rbSuggested.CheckedChanged += rbSuggested_CheckedChanged;
            rbManual.CheckedChanged += rbManual_CheckedChanged;

            LoadSourceTypes();
        }

        private void LoadSourceTypes()
        {
            try
            {
                using (var db = new MydbDB())
                {
                    // 取得目標料號的PTFE資訊
                    var targetTypeInfo = db.Types.FirstOrDefault(t => t.TypeColumn == targetType);
                    string targetPTFE = targetTypeInfo?.PTFEColor ?? "";

                    // 修改：自動建議包含目標料號自己（允許參考自己）
                    var suggestedTypes = db.Types
                        .Where(t => !string.IsNullOrEmpty(t.PTFEColor) && t.PTFEColor == targetPTFE)
                        .Select(t => new { t.TypeColumn, t.PTFEColor })
                        .OrderByDescending(t => t.TypeColumn == targetType) // 目標料號排在最前面
                        .ToList();

                    // 修改：所有可用的料號包含目標料號自己
                    var allTypes = db.Types
                        .Select(t => new { t.TypeColumn, t.PTFEColor })
                        .OrderByDescending(t => t.TypeColumn == targetType) // 目標料號排在最前面
                        .ToList();

                    // 建立建議區域的 RadioButton
                    if (suggestedTypes.Any())
                    {
                        lblSuggestion.Text = $"相同PTFE({targetPTFE})的現有料號:";

                        int yPosition = 10;

                        foreach (var type in suggestedTypes)
                        {
                            var rb = new RadioButton();
                            // 標註當前料號
                            string displayText = type.TypeColumn == targetType
                                ? $"{type.TypeColumn} ({type.PTFEColor}) - 當前料號"
                                : $"{type.TypeColumn} ({type.PTFEColor})";
                            rb.Text = displayText;
                            rb.Tag = type.TypeColumn;
                            rb.AutoSize = true;
                            rb.Location = new System.Drawing.Point(20, yPosition);

                            // 訂閱 CheckedChanged 事件
                            rb.CheckedChanged += SuggestedRadioButton_CheckedChanged;

                            panelSuggested.Controls.Add(rb);

                            yPosition += 25;
                        }

                        // 預設選中「使用建議料號」
                        rbSuggested.Checked = true;
                    }
                    else
                    {
                        lblSuggestion.Text = "沒有找到相同PTFE的料號";

                        // 沒有建議時，強制選擇手動模式
                        rbManual.Checked = true;
                        rbSuggested.Enabled = false;
                    }

                    // 填入手動選擇下拉選單（標註當前料號）
                    cmbAllTypes.DisplayMember = "Display";
                    cmbAllTypes.ValueMember = "type";
                    cmbAllTypes.DataSource = allTypes.Select(t => new
                    {
                        type = t.TypeColumn,
                        Display = t.TypeColumn == targetType
                            ? $"{t.TypeColumn} ({t.PTFEColor}) - 當前料號"
                            : $"{t.TypeColumn} ({t.PTFEColor})"
                    }).ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入料號資料失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 建議區域 RadioButton 的事件處理
        private void SuggestedRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            var rb = sender as RadioButton;
            if (rb != null && rb.Checked)
            {
                selectedSuggestedType = rb.Tag.ToString();

                // 自動切換到「使用建議料號」
                if (!rbSuggested.Checked)
                {
                    rbSuggested.Checked = true;
                }
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (rbSuggested.Checked)
            {
                // 使用建議料號
                if (!string.IsNullOrEmpty(selectedSuggestedType))
                {
                    SelectedSourceType = selectedSuggestedType;
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("請選擇一個建議的料號", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else if (rbManual.Checked)
            {
                // 使用手動選擇
                if (cmbAllTypes.SelectedValue != null)
                {
                    SelectedSourceType = cmbAllTypes.SelectedValue.ToString();
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("請選擇一個料號", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("請選擇使用建議或手動選擇", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        // 切換到建議模式
        private void rbSuggested_CheckedChanged(object sender, EventArgs e)
        {
            if (rbSuggested.Checked)
            {
                // 啟用建議區域，停用手動選擇
                panelSuggested.Enabled = true;
                cmbAllTypes.Enabled = false;

                // 確保建議區域有選中的項目
                var checkedRb = panelSuggested.Controls.OfType<RadioButton>().FirstOrDefault(rb => rb.Checked);
                if (checkedRb == null)
                {
                    var firstRb = panelSuggested.Controls.OfType<RadioButton>().FirstOrDefault();
                    if (firstRb != null)
                    {
                        firstRb.Checked = true;
                        selectedSuggestedType = firstRb.Tag.ToString();
                    }
                }
            }
            else
            {
                // 停用建議區域，清除選擇
                panelSuggested.Enabled = false;

                // 清除建議區域的所有選擇
                foreach (var rb in panelSuggested.Controls.OfType<RadioButton>())
                {
                    rb.Checked = false;
                }
                selectedSuggestedType = null;
            }
        }

        // 切換到手動模式
        private void rbManual_CheckedChanged(object sender, EventArgs e)
        {
            if (rbManual.Checked)
            {
                // 啟用手動選擇，停用建議區域
                cmbAllTypes.Enabled = true;
                panelSuggested.Enabled = false;

                // 清除建議區域的所有選擇
                foreach (var rb in panelSuggested.Controls.OfType<RadioButton>())
                {
                    rb.Checked = false;
                }
                selectedSuggestedType = null;
            }
            else
            {
                // 停用手動選擇
                cmbAllTypes.Enabled = false;
            }
        }
    }
}