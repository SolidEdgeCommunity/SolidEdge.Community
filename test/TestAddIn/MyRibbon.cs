using SolidEdge.Community.AddIn;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestAddIn
{
    //[RibbonAttribute(SolidEdge.CATID.SEApplication)]
    //[RibbonAttribute(SolidEdge.CATID.SEPart)]
    //[RibbonAttribute(SolidEdge.CATID.SEDMPart)]
    public class MyRibbon : Ribbon
    {
        const string _embeddedResourceName = "TestAddIn.Ribbon.xml";
        private RibbonControl _buttonOpenGlBoxes;

        public MyRibbon()
            : base()
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();

            this.LoadXml(assembly, "TestAddIn.Ribbon.xml");

            _buttonOpenGlBoxes = GetButton(21);
            _buttonOpenGlBoxes.Click += _buttonOpenGlBoxes_Click;
        }

        void _buttonOpenGlBoxes_Click(RibbonControl control)
        {
        }
    }
}
