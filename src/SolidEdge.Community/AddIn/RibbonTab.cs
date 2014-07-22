using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.AddIn
{
    public class RibbonTab
    {
        private Ribbon _ribbon;
        private string _name;
        private List<RibbonGroup> _groups = new List<RibbonGroup>();

        internal RibbonTab(Ribbon ribbon, string name)
        {
            _ribbon = ribbon;
            _name = name;
        }

        public RibbonGroup AddGroup(string name)
        {
            var ribbonGroup = new RibbonGroup(this, name);
            _groups.Add(ribbonGroup);
            return ribbonGroup;
        }

        public Ribbon Ribbon { get { return _ribbon; } }
        public string Name { get { return _name; } }

        public System.Collections.Generic.IEnumerable<RibbonButton> Buttons
        {
            get
            {
                foreach (var control in this.Controls.OfType<RibbonButton>())
                {
                    yield return control;
                }
            }
        }

        public System.Collections.Generic.IEnumerable<RibbonCheckBox> CheckBoxes
        {
            get
            {
                foreach (var control in this.Controls.OfType<RibbonCheckBox>())
                {
                    yield return control;
                }
            }
        }

        public System.Collections.Generic.IEnumerable<RibbonControl> Controls
        {
            get
            {

                foreach (var group in this.Groups)
                {
                    foreach (var control in group.Controls)
                    {
                        yield return control;
                    }
                }
            }
        }

        public System.Collections.Generic.IEnumerable<RibbonGroup> Groups
        {
            get
            {
                foreach (var group in _groups)
                {
                    yield return group;
                }
            }
        }

        public System.Collections.Generic.IEnumerable<RibbonRadioButton> RadioButtons
        {
            get
            {
                foreach (var control in this.Controls.OfType<RibbonRadioButton>())
                {
                    yield return control;
                }
            }
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
