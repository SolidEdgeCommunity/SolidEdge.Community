using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeFramework.Extensions
{
    /// <summary>
    /// SolidEdgeFramework.Documents extension methods.
    /// </summary>
    public static class DocumentsExtensions
    {
        /// <summary>
        /// Creates a new document.
        /// </summary>
        internal static T Add<T>(this SolidEdgeFramework.Documents documents, string progId) where T : class
        {
            return (T)documents.Add(progId);
        }

        /// <summary>
        /// Creates a new document.
        /// </summary>
        internal static T Add<T>(this SolidEdgeFramework.Documents documents, string progId, object TemplateDoc) where T : class
        {
            return (T)documents.Add(progId, TemplateDoc);
        }

        /// <summary>
        /// Creates a new assembly document.
        /// </summary>
        public static SolidEdgeAssembly.AssemblyDocument AddAssemblyDocument(this SolidEdgeFramework.Documents documents)
        {
            return documents.Add<SolidEdgeAssembly.AssemblyDocument>(SolidEdge.PROGID.AssemblyDocument);
        }

        /// <summary>
        /// Creates a new assembly document.
        /// </summary>
        public static SolidEdgeAssembly.AssemblyDocument AddAssemblyDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgeAssembly.AssemblyDocument>(SolidEdge.PROGID.AssemblyDocument, TemplateDoc);
        }

        /// <summary>
        /// Creates a new draft document.
        /// </summary>
        public static SolidEdgeDraft.DraftDocument AddDraftDocument(this SolidEdgeFramework.Documents documents)
        {
            return documents.Add<SolidEdgeDraft.DraftDocument>(SolidEdge.PROGID.DraftDocument);
        }

        /// <summary>
        /// Creates a new draft document.
        /// </summary>
        public static SolidEdgeDraft.DraftDocument AddDraftDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgeDraft.DraftDocument>(SolidEdge.PROGID.DraftDocument, TemplateDoc);
        }

        /// <summary>
        /// Creates a new part document.
        /// </summary>
        public static SolidEdgePart.PartDocument AddPartDocument(this SolidEdgeFramework.Documents documents)
        {
            return documents.Add<SolidEdgePart.PartDocument>(SolidEdge.PROGID.PartDocument);
        }

        /// <summary>
        /// Creates a new part document.
        /// </summary>
        public static SolidEdgePart.PartDocument AddPartDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgePart.PartDocument>(SolidEdge.PROGID.PartDocument, TemplateDoc);
        }

        /// <summary>
        /// Creates a new sheetmetal document.
        /// </summary>
        public static SolidEdgePart.SheetMetalDocument AddSheetMetalDocument(this SolidEdgeFramework.Documents documents)
        {
            return documents.Add<SolidEdgePart.SheetMetalDocument>(SolidEdge.PROGID.SheetMetalDocument);
        }

        /// <summary>
        /// Creates a new sheetmetal document.
        /// </summary>
        public static SolidEdgePart.SheetMetalDocument AddSheetMetalDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgePart.SheetMetalDocument>(SolidEdge.PROGID.SheetMetalDocument, TemplateDoc);
        }
    }
}
