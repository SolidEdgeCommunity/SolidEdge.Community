using SolidEdgeCommunity;
using SolidEdgeCommunity.Extensions;
using System;
using System.Collections.Generic;
using System.Drawing;
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
            try
            {
                var application = SolidEdgeUtils.Connect();
                var draftDocument = application.GetActiveDocument<SolidEdgeDraft.DraftDocument>();

                foreach (var drawingObject in draftDocument.ActiveSection.EnumerateDrawingObjects())
                {
                    var type = SolidEdgeCommunity.Runtime.InteropServices.ComObject.GetType(drawingObject);
                    Console.WriteLine(drawingObject);
                }
                
                //using (var task = new IsolatedTask<MyIsolatedTask>())
                //{
                //    task.Proxy.Application = null;// application;
                //    task.Proxy.Document = application.GetActiveDocument();
                //    var results = task.Proxy.DoWork("Hello world!");
                //}
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
