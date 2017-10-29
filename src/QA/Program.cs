using SolidEdgeCommunity;
using SolidEdgeCommunity.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QA
{
    class Program
    {
        static void Main(string[] args)
        {
            var application = SolidEdgeUtils.Connect();
            var documents = application.Documents;
            var assemblyDocument = documents.AddAssemblyDocument();
            var draftDocument = documents.AddDraftDocument();
        }
    }
}
