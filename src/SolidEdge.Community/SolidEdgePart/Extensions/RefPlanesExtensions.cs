using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgePart.Extensions
{
    /// <summary>
    /// SolidEdgePart.RefPlanes extension methods.
    /// </summary>
    public static class RefPlanesExtensions
    {
        /// <summary>
        /// Returns a SolidEdgePart.RefPlane representing the default top plane.
        /// </summary>
        public static SolidEdgePart.RefPlane GetTopPlane(this SolidEdgePart.RefPlanes refPlanes)
        {
            return refPlanes.Item(1);
        }

        /// <summary>
        /// Returns a SolidEdgePart.RefPlane representing the default right plane.
        /// </summary>
        public static SolidEdgePart.RefPlane GetRightPlane(this SolidEdgePart.RefPlanes refPlanes)
        {
            return refPlanes.Item(2);
        }

        /// <summary>
        /// Returns a SolidEdgePart.RefPlane representing the default front plane.
        /// </summary>
        public static SolidEdgePart.RefPlane GetFrontPlane(this SolidEdgePart.RefPlanes refPlanes)
        {
            return refPlanes.Item(3);
        }
    }
}
