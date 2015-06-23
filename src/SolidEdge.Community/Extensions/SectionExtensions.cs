using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.Extensions
{
    /// <summary>
    /// SolidEdgeDraft.Section extensions.
    /// </summary>
    public static class SectionExtensions
    {
        /// <summary>
        /// Returns an enumerable collection of drawing objects.
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
        public static IEnumerable<object> EnumerateDrawingObjects(this SolidEdgeDraft.Section section)
        {
            foreach (var drawingObject in EnumerateDrawingObjects<object>(section))
            {
                yield return drawingObject;
            }
        }

        /// <summary>
        /// Returns an enumerable collection of drawing objects of the specified type.
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public static IEnumerable<T> EnumerateDrawingObjects<T>(this SolidEdgeDraft.Section section) where T : class
        {
            foreach (SolidEdgeDraft.Sheet sheet in section.Sheets)
            {
                // The following section types are hidden and should not be used.
                if (sheet.SectionType == SolidEdgeDraft.SheetSectionTypeConstants.igDrawingViewSection) continue;
                if (sheet.SectionType == SolidEdgeDraft.SheetSectionTypeConstants.igBlockViewSection) continue;

                // Should work but throws an exception...
                //foreach (var drawingObject in sheet.DrawingObjects)

                var drawingObjects = sheet.DrawingObjects;

                for (int i = 1; i <= drawingObjects.Count; i++)
                {
                    var drawingObject = drawingObjects.Item(i);

                    if (drawingObject is T)
                    {
                        yield return (T)drawingObject;
                    }
                }
            }
        }

        /// <summary>
        /// Returns an enumerable collection of drawing views.
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
        public static IEnumerable<SolidEdgeDraft.DrawingView> EnumerateDrawingViews(this SolidEdgeDraft.Section section)
        {
            foreach (SolidEdgeDraft.Sheet sheet in section.Sheets)
            {
                foreach (SolidEdgeDraft.DrawingView drawingView in sheet.DrawingViews)
                {
                    yield return drawingView;
                }
            }
        }
    }
}
