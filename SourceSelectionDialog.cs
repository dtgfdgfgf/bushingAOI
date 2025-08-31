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

        public SourceSelectionDialog(string targetType)
        {
            InitializeComponent();
            this.targetType = targetType;
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

                    // 自動建議：相同PTFE的料號
                    var suggestedTypes = db.Types
                        .Where(t => t.TypeColumn != targetType && !string.IsNullOrEmpty(t.PTFEColor) && t.PTFEColor == targetPTFE)
                        .Select(t => new { t.TypeColumn, t.PTFEColor })
                        .ToList();

                    // 所有可用的料號
                    var allTypes = db.Types
                        .Where(t => t.TypeColumn != targetType)
                        .Select(t => new { t.TypeColumn, t.PTFEColor })
                        .ToList();

                    // 填入自動建議
                    if (suggestedTypes.Any())
                    {
                        lblSuggestion.Text = $"相同PTFE({targetPTFE})的現有料號:";
                        foreach (var type in suggestedTypes)
                        {
                            var rb = new RadioButton();
                            rb.Text = $"{type.TypeColumn} ({type.PTFEColor})";
                            rb.Tag = type.TypeColumn;
                            rb.AutoSize = true;
                            rb.Location = new Point(20, panelSuggested.Controls.Count * 25 + 10);
                            panelSuggested.Controls.Add(rb);

                            if (panelSuggested.Controls.Count == 1)
                            {
                                rb.Checked = true; // 預選第一個
                            }
                        }
                    }
                    else
                    {
                        lblSuggestion.Text = "沒有找到相同PTFE的料號";
                    }

                    // 填入手動選擇下拉選單
                    cmbAllTypes.DisplayMember = "Display";
                    cmbAllTypes.ValueMember = "type";
                    cmbAllTypes.DataSource = allTypes.Select(t => new 
                    { 
                        type = t.TypeColumn, 
                        Display = $"{t.TypeColumn} ({t.PTFEColor})" 
                    }).ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入料號資料失敗: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (rbSuggested.Checked)
            {
                // 使用建議的料號
                var selectedRadio = panelSuggested.Controls.OfType<RadioButton>().FirstOrDefault(rb => rb.Checked);
                if (selectedRadio != null)
                {
                    SelectedSourceType = selectedRadio.Tag.ToString();
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
                // 使用手動選擇的料號
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

        private void rbSuggested_CheckedChanged(object sender, EventArgs e)
        {
            panelSuggested.Enabled = rbSuggested.Checked;
            cmbAllTypes.Enabled = !rbSuggested.Checked;
        }

        private void rbManual_CheckedChanged(object sender, EventArgs e)
        {
            cmbAllTypes.Enabled = rbManual.Checked;
            panelSuggested.Enabled = !rbManual.Checked;
        }
    }
}
