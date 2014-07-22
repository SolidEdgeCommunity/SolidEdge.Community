using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SolidEdgeCommunity.AddIn
{
    public class EdgeBarControl : UserControl
    {
        private EdgeBarPage _edgeBarPage;
        private string _tooltip = String.Empty;

        public event EventHandler AfterInitialize;

        #region Methods

        public void Initialize(EdgeBarPage edgeBarPage)
        {
            if (edgeBarPage == null) throw new ArgumentNullException("edgeBarPage");
            if (_edgeBarPage != null) { throw new InvalidOperationException("Control has already been initialized."); }

            _edgeBarPage = edgeBarPage;

            if (AfterInitialize != null)
            {
                AfterInitialize(this, null);
            }
        }

        #endregion

        #region Properties

        [Browsable(false)]
        public EdgeBarPage EdgeBarPage
        {
            get { return _edgeBarPage; }
        }

        [Browsable(false)]
        public SolidEdgeFramework.SolidEdgeDocument Document
        {
            get { return _edgeBarPage.Document; }
        }

        /// <summary>
        /// The string to be used in the EdgeBarPage caption and tooltip.
        /// </summary>
        [Browsable(true)]
        public string ToolTip
        {
            get { return _tooltip; }
            set { _tooltip = value; }
        }

        #endregion
    }
}
