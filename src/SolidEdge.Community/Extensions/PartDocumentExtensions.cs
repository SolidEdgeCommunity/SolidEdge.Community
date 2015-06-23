using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.Extensions
{
    /// <summary>
    /// SolidEdgePart.PartDocument extension methods.
    /// </summary>
    public static class PartDocumentExtensions
    {
        /// <summary>
        /// Returns the version of Solid Edge that was used to create the referenced document.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static Version GetCreatedVersion(this SolidEdgePart.PartDocument document)
        {
            return new Version(document.CreatedVersion);
        }

        /// <summary>
        /// Returns the version of Solid Edge that was used the last time the referenced document was saved.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static Version GetLastSavedVersion(this SolidEdgePart.PartDocument document)
        {
            return new Version(document.LastSavedVersion);
        }

        /// <summary>
        /// Returns the properties for the referenced document.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static SolidEdgeFramework.PropertySets GetProperties(this SolidEdgePart.PartDocument document)
        {
            return document.Properties as SolidEdgeFramework.PropertySets;
        }

        /// <summary>
        /// Returns the summary information property set for the referenced document.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static SolidEdgeFramework.SummaryInfo GetSummaryInfo(this SolidEdgePart.PartDocument document)
        {
            return document.SummaryInfo as SolidEdgeFramework.SummaryInfo;
        }

        /// <summary>
        /// Returns a collection of variables for the referenced document.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static SolidEdgeFramework.Variables GetVariables(this SolidEdgePart.PartDocument document)
        {
            return document.Variables as SolidEdgeFramework.Variables;
        }
    }
}
