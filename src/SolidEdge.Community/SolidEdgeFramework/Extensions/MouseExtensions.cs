using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeFramework.Extensions
{
    public static class MouseExtensions
    {
        public static void AddToLocateFilter(this SolidEdgeFramework.Mouse mouse, SolidEdgeConstants.seLocateFilterConstants filter)
        {
            mouse.AddToLocateFilter((int)filter);
        }

        public static void SetLocateMode(this SolidEdgeFramework.Mouse mouse, SolidEdgeConstants.seLocateModes mode)
        {
            mouse.LocateMode = (int)mode;
        }

        public static SolidEdgeConstants.seLocateModes GetLocateMode(this SolidEdgeFramework.Mouse mouse, SolidEdgeConstants.seLocateModes mode)
        {
            return (SolidEdgeConstants.seLocateModes)mouse.LocateMode;
        }
    }
}
