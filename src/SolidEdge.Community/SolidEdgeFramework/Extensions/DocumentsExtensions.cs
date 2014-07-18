using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeFramework.Extensions
{
    public static class DocumentsExtensions
    {
        internal static T Add<T>(this SolidEdgeFramework.Documents documents, string progId) where T : class
        {
            return (T)documents.Add(progId);
        }

        internal static T Add<T>(this SolidEdgeFramework.Documents documents, string progId, object TemplateDoc) where T : class
        {
            return (T)documents.Add(progId, TemplateDoc);
        }

        public static SolidEdgeAssembly.AssemblyDocument AddAssemblyDocument(this SolidEdgeFramework.Documents documents)
        {
            return documents.Add<SolidEdgeAssembly.AssemblyDocument>(SolidEdge.PROGID.AssemblyDocument);
        }

        public static SolidEdgeAssembly.AssemblyDocument AddAssemblyDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgeAssembly.AssemblyDocument>(SolidEdge.PROGID.AssemblyDocument, TemplateDoc);
        }

        public static SolidEdgeDraft.DraftDocument AddDraftDocument(this SolidEdgeFramework.Documents documents)
        {
            return documents.Add<SolidEdgeDraft.DraftDocument>(SolidEdge.PROGID.DraftDocument);
        }

        public static SolidEdgeDraft.DraftDocument AddDraftDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgeDraft.DraftDocument>(SolidEdge.PROGID.DraftDocument, TemplateDoc);
        }

        public static SolidEdgePart.PartDocument AddPartDocument(this SolidEdgeFramework.Documents documents)
        {
            return documents.Add<SolidEdgePart.PartDocument>(SolidEdge.PROGID.PartDocument);
        }

        public static SolidEdgePart.PartDocument AddPartDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgePart.PartDocument>(SolidEdge.PROGID.PartDocument, TemplateDoc);
        }

        public static SolidEdgePart.SheetMetalDocument AddSheetMetalDocument(this SolidEdgeFramework.Documents documents)
        {
            return documents.Add<SolidEdgePart.SheetMetalDocument>(SolidEdge.PROGID.SheetMetalDocument);
        }

        public static SolidEdgePart.SheetMetalDocument AddSheetMetalDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgePart.SheetMetalDocument>(SolidEdge.PROGID.SheetMetalDocument, TemplateDoc);
        }
    }
}
