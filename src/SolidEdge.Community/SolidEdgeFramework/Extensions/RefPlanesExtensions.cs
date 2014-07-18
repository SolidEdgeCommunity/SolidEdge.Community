using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeFramework.Extensions
{
    public static class RefPlanesExtensions
    {
        public static SolidEdgePart.RefPlane GetTopPlane(this SolidEdgePart.RefPlanes refPlanes)
        {
            return refPlanes.Item(1);
        }

        public static SolidEdgePart.RefPlane GetRightPlane(this SolidEdgePart.RefPlanes refPlanes)
        {
            return refPlanes.Item(2);
        }

        public static SolidEdgePart.RefPlane GetFrontPlane(this SolidEdgePart.RefPlanes refPlanes)
        {
            return refPlanes.Item(3);
        }
    }
}
