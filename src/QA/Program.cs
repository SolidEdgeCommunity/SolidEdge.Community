using SolidEdgeCommunity;
using System;

namespace QA
{
    class Program
    {
        static void Main(string[] args)
        {
            var application = SolidEdgeUtils.Connect(true);
            application.Visible = true;
        }
    }
}
