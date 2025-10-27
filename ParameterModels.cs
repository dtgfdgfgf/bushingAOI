using System;
using System.Collections.Generic;

namespace peilin
{
    public class ParameterItem
    {
        public string Type { get; set; }
        public string Name { get; set; }
        public string Value { get; set; }
        public int? Stop { get; set; }
        public string ChineseName { get; set; }
        public ParameterZone Zone { get; set; }
        public bool IsSelected { get; set; }
    }

    public enum ParameterZone
    {
        Reference,         // 參考區（來源料號參數，只讀）
        AddedUnmodified,   // 已新增未修改區
        AddedModified      // 已新增已修改區
    }

    public class ParameterSession
    {
        public string SessionId { get; set; }
        public DateTime CreatedTime { get; set; }
        public string SourceType { get; set; }
        public string TargetType { get; set; }
        public List<ParameterItem> Parameters { get; set; }

        public ParameterSession()
        {
            Parameters = new List<ParameterItem>();
        }
    }
}
