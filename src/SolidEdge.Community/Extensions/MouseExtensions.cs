using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.Extensions
{
    /// <summary>
    /// SolidEdgeFramework.Mouse extension methods.
    /// </summary>
    public static class MouseExtensions
    {
        /// <summary>
        /// Specifies what types of objects the Mouse can locate.
        /// </summary>
        /// <param name="mouse"></param>
        /// <param name="filter"></param>
        public static void AddToLocateFilter(this SolidEdgeFramework.Mouse mouse, SolidEdgeConstants.seLocateFilterConstants filter)
        {
            mouse.AddToLocateFilter((int)filter);
        }

        /// <summary>
        /// Sets the locate mode for the referenced object.
        /// </summary>
        /// <param name="mouse"></param>
        /// <param name="mode"></param>
        public static void SetLocateMode(this SolidEdgeFramework.Mouse mouse, SolidEdgeConstants.seLocateModes mode)
        {
            mouse.LocateMode = (int)mode;
        }

        /// <summary>
        /// Returns the locate mode for the referenced object.
        /// </summary>
        /// <param name="mouse"></param>
        /// <param name="mode"></param>
        /// <returns></returns>
        public static SolidEdgeConstants.seLocateModes GetLocateMode(this SolidEdgeFramework.Mouse mouse, SolidEdgeConstants.seLocateModes mode)
        {
            return (SolidEdgeConstants.seLocateModes)mouse.LocateMode;
        }
    }
}
