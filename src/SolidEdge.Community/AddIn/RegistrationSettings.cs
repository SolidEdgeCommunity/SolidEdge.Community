using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.AddIn
{
    /// <summary>
    /// 
    /// </summary>
    public sealed class RegistrationSettings
    {
        private List<Guid> _implementedCategories = new List<Guid>();
        private List<Guid> _environments = new List<Guid>();
        private Dictionary<CultureInfo, string> _titles = new Dictionary<CultureInfo, string>();
        private Dictionary<CultureInfo, string> _summaries = new Dictionary<CultureInfo, string>();

        public RegistrationSettings()
            : this(null)
        {
        }

        public RegistrationSettings(Type type)
        {
            this.Type = type;
            Enabled = true;
            _implementedCategories.Add(new Guid(SolidEdgeSDK.CATID.SolidEdgeAddIn));
        }

        public Type Type { get; set; }
        public bool Enabled { get; set; }

        public Guid[] ImplementedCategories { get { return _implementedCategories.ToArray(); } }
        public List<Guid> Environments { get { return _environments; } }
        public Dictionary<CultureInfo, string> Titles { get { return _titles; } }
        public Dictionary<CultureInfo, string> Summaries { get { return _summaries; } }
    }
}
