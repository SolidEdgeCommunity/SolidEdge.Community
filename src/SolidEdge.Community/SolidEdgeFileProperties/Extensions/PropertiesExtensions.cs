using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeFileProperties.Extensions
{
    /// <summary>
    /// SolidEdgeFileProperties.Properties extensions.
    /// </summary>
    public static class PropertiesExtensions
    {
        /// <summary>
        /// Returns all properties of a given property set as an IEnumerable.
        /// </summary>
        /// <param name="properties"></param>
        public static IEnumerable<SolidEdgeFileProperties.Property> AsEnumerable(this SolidEdgeFileProperties.Properties properties)
        {
            List<SolidEdgeFileProperties.Property> list = new List<SolidEdgeFileProperties.Property>();

            foreach (var property in properties)
            {
                list.Add((SolidEdgeFileProperties.Property)property);
            }

            return list.AsEnumerable();
        }
    }
}
