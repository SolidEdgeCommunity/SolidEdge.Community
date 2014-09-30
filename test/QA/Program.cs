using SolidEdgeCommunity.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace QA
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var filename = @"C:\Program Files\Solid Edge ST7\Training\Coffee Pot.asm";
            var application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
            var documents = application.Documents;

            var assembly = documents.Add<SolidEdgeAssembly.AssemblyDocument>();
            var draft = documents.Add<SolidEdgeDraft.DraftDocument>();
            var part = documents.Add<SolidEdgePart.PartDocument>();
            var sheetmetal = documents.Add<SolidEdgePart.SheetMetalDocument>();
        }
    }
}
