using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
        internal static TDocumentType Add<TDocumentType>(this SolidEdgeFramework.Documents documents, string progId) where TDocumentType : class
        {
            return (TDocumentType)documents.Add(progId);
        }

        /// <summary>
        /// Creates a new document.
        /// </summary>
        internal static TDocumentType Add<TDocumentType>(this SolidEdgeFramework.Documents documents, string progId, object TemplateDoc) where TDocumentType : class
        {
            return (TDocumentType)documents.Add(progId, TemplateDoc);
        }

        /// <summary>
        /// Creates a new document.
        /// </summary>
        /// <typeparam name="TDocumentType">SolidEdgeAssembly.AssemblyDocument, SolidEdgeDraft.DraftDocument, SolidEdgePart.PartDocument or SolidEdgePart.SheetMetalDocument.</typeparam>
        /// <param name="documents"></param>
        /// <returns></returns>
        public static TDocumentType Add<TDocumentType>(this SolidEdgeFramework.Documents documents) where TDocumentType : class
        {
            var t = typeof(TDocumentType);

            if (t.Equals(typeof(SolidEdgeAssembly.AssemblyDocument)))
            {
                return (TDocumentType)documents.AddAssemblyDocument();
            }
            else if (t.Equals(typeof(SolidEdgeDraft.DraftDocument)))
            {
                return (TDocumentType)documents.AddDraftDocument();
            }
            else if (t.Equals(typeof(SolidEdgePart.PartDocument)))
            {
                return (TDocumentType)documents.AddPartDocument();
            }
            else if (t.Equals(typeof(SolidEdgePart.SheetMetalDocument)))
            {
                return (TDocumentType)documents.AddSheetMetalDocument();
            }
            else
            {
                throw new NotSupportedException();
            }
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

        /// <summary>
        /// Opens an existing Solid Edge document.
        /// </summary>
        public static TDocumentType Open<TDocumentType>(this SolidEdgeFramework.Documents documents, string Filename) where TDocumentType : class
        {
            return (TDocumentType)documents.Open(Filename);
        }

        /// <summary>
        /// Opens an existing Solid Edge document.
        /// </summary>
        public static TDocumentType Open<TDocumentType>(this SolidEdgeFramework.Documents documents, string Filename, object DocRelationAutoServer, object AltPath, object RecognizeFeaturesIfPartTemplate, object RevisionRuleOption, object StopFileOpenIfRevisionRuleNotApplicable)
        {
            return (TDocumentType)documents.Open(Filename, DocRelationAutoServer, AltPath, RecognizeFeaturesIfPartTemplate, RevisionRuleOption, StopFileOpenIfRevisionRuleNotApplicable);
        }

        /// <summary>
        /// Opens an existing Solid Edge document in the background with no window.
        /// </summary>
        public static TDocumentType OpenInBackground<TDocumentType>(this SolidEdgeFramework.Documents documents, string Filename)
        {
            ulong JDOCUMENTPROP_NOWINDOW = 0x00000008;
            return (TDocumentType)documents.Open(Filename, JDOCUMENTPROP_NOWINDOW);
        }

        /// <summary>
        /// Opens an existing Solid Edge document in the background with no window.
        /// </summary>
        public static TDocumentType OpenInBackground<TDocumentType>(this SolidEdgeFramework.Documents documents, string Filename, object AltPath, object RecognizeFeaturesIfPartTemplate, object RevisionRuleOption, object StopFileOpenIfRevisionRuleNotApplicable)
        {
            ulong JDOCUMENTPROP_NOWINDOW = 0x00000008;
            return (TDocumentType)documents.Open(Filename, JDOCUMENTPROP_NOWINDOW, AltPath, RecognizeFeaturesIfPartTemplate, RevisionRuleOption, StopFileOpenIfRevisionRuleNotApplicable);
        }
    }
}
