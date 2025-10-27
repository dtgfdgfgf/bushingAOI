using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace peilin
{
    public partial class SaveConfirmDialog : Form
    {
        private List<ParameterItem> parameters;
        private string targetType;

        public SaveConfirmDialog(List<ParameterItem> parameters, string targetType)
        {
            InitializeComponent();
            this.parameters = parameters;
            this.targetType = targetType;
            LoadParameterList();
        }

        private void LoadParameterList()
        {
            lblInfo.Text = $"即將新增 {parameters.Count} 個參數到料號: {targetType}";

            // 按參數名排序顯示
            var sortedParams = parameters.OrderBy(p => p.Name).ThenBy(p => p.Stop).ToList();

            dgvParameters.DataSource = sortedParams.Select(p => new
            {
                參數名 = p.Name,
                站點 = p.Stop,
                參數值 = p.Value,
                中文名稱 = p.ChineseName,
                狀態 = GetZoneName(p.Zone)
            }).ToList();

            // 設定欄位寬度
            if (dgvParameters.Columns.Count > 0)
            {
                dgvParameters.Columns["參數名"].Width = 150;
                dgvParameters.Columns["站點"].Width = 60;
                dgvParameters.Columns["參數值"].Width = 100;
                dgvParameters.Columns["中文名稱"].Width = 120;
                dgvParameters.Columns["狀態"].Width = 80;
            }

            dgvParameters.ReadOnly = true;
            dgvParameters.AllowUserToAddRows = false;
            dgvParameters.AllowUserToDeleteRows = false;
        }

        // 由 GitHub Copilot 產生
        private string GetZoneName(ParameterZone zone)
        {
            switch (zone)
            {
                case ParameterZone.Reference: return "參考區";
                case ParameterZone.AddedUnmodified: return "已新增未修改";
                case ParameterZone.AddedModified: return "已新增已修改";
                default: return "未知";
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
