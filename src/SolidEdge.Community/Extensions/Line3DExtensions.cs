using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.Extensions
{
    /// <summary>
    /// SolidEdgePart.Line3D extension methods.
    /// </summary>
    public static class Line3DExtensions
    {
        public static double[] SafeGetKeypointPosition(this SolidEdgePart.Line3D line3d, SolidEdgePart.Sketch3DKeypointType KeypointType, bool throwOnError = false)
        {
            var position = Array.CreateInstance(typeof(double), 0);

            try
            {
                // GetKeypointPosition may throw an exception so wrap in try\catch.
                line3d.GetKeypointPosition(KeypointType, ref position);
            }
            catch
            {
                if (throwOnError) throw;
            }

            return position.OfType<double>().ToArray();
        }
    }
}
