using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.Extensions
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
            return documents.Add<SolidEdgeAssembly.AssemblyDocument>(SolidEdgeSDK.PROGID.SolidEdge_AssemblyDocument);
        }

        /// <summary>
        /// Creates a new assembly document.
        /// </summary>
        public static SolidEdgeAssembly.AssemblyDocument AddAssemblyDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgeAssembly.AssemblyDocument>(SolidEdgeSDK.PROGID.SolidEdge_AssemblyDocument, TemplateDoc);
        }

        /// <summary>
        /// Creates a new draft document.
        /// </summary>
        public static SolidEdgeDraft.DraftDocument AddDraftDocument(this SolidEdgeFramework.Documents documents)
        {
            return documents.Add<SolidEdgeDraft.DraftDocument>(SolidEdgeSDK.PROGID.SolidEdge_DraftDocument);
        }

        /// <summary>
        /// Creates a new draft document.
        /// </summary>
        public static SolidEdgeDraft.DraftDocument AddDraftDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgeDraft.DraftDocument>(SolidEdgeSDK.PROGID.SolidEdge_DraftDocument, TemplateDoc);
        }

        /// <summary>
        /// Creates a new part document.
        /// </summary>
        public static SolidEdgePart.PartDocument AddPartDocument(this SolidEdgeFramework.Documents documents)
        {
            return documents.Add<SolidEdgePart.PartDocument>(SolidEdgeSDK.PROGID.SolidEdge_PartDocument);
        }

        /// <summary>
        /// Creates a new part document.
        /// </summary>
        public static SolidEdgePart.PartDocument AddPartDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgePart.PartDocument>(SolidEdgeSDK.PROGID.SolidEdge_PartDocument, TemplateDoc);
        }

        /// <summary>
        /// Creates a new sheetmetal document.
        /// </summary>
        public static SolidEdgePart.SheetMetalDocument AddSheetMetalDocument(this SolidEdgeFramework.Documents documents)
        {
            return documents.Add<SolidEdgePart.SheetMetalDocument>(SolidEdgeSDK.PROGID.SolidEdge_SheetMetalDocument);
        }

        /// <summary>
        /// Creates a new sheetmetal document.
        /// </summary>
        public static SolidEdgePart.SheetMetalDocument AddSheetMetalDocument(this SolidEdgeFramework.Documents documents, object TemplateDoc)
        {
            return documents.Add<SolidEdgePart.SheetMetalDocument>(SolidEdgeSDK.PROGID.SolidEdge_SheetMetalDocument, TemplateDoc);
        }
    }
}
