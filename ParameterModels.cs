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
        Unmodified,
        Modified,
        Fixed
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
