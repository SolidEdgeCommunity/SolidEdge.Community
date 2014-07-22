using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.AddIn
{
    public class RibbonGroup
    {
        private RibbonTab _tab;
        private string _name = String.Empty;
        private List<RibbonControl> _controls = new List<RibbonControl>();

        internal RibbonGroup(RibbonTab tab, string name)
        {
            _tab = tab;
            _name = name;
        }

        public void AddControl(RibbonControl control)
        {
            control.Group = this;
            _controls.Add(control);
        }

        public RibbonTab Tab { get { return _tab; } }
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
                foreach (var control in _controls)
                {
                    yield return control;
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
