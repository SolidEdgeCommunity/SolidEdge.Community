using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.Extensions
{
    /// <summary>
    /// SolidEdgeFileProperties.PropertySets extensions.
    /// </summary>
    public static class PropertySetsExtensions
    {
        ///// <summary>
        ///// Returns all property sets as an IEnumerable.
        ///// </summary>
        ///// <param name="propertySets"></param>
        //public static IEnumerable<SolidEdgeFileProperties.Properties> AsEnumerable(this SolidEdgeFileProperties.PropertySets propertySets)
        //{
        //    List<SolidEdgeFileProperties.Properties> list = new List<SolidEdgeFileProperties.Properties>();

        //    foreach (var properties in propertySets)
        //    {
        //        list.Add((SolidEdgeFileProperties.Properties)properties);
        //    }

        //    return list.AsEnumerable();
        //}

        /// <summary>
        /// Closes an open document with the option to save changes.
        /// </summary>
        /// <param name="propertySets"></param>
        /// <param name="saveChanges"></param>
        public static void Close(this SolidEdgeFileProperties.PropertySets propertySets, bool saveChanges)
        {
            if (saveChanges)
            {
                propertySets.Save();
            }

            propertySets.Close();
        }

        /// <summary>
        /// Adds or modifies a custom property.
        /// </summary>
        internal static SolidEdgeFileProperties.Property InternalUpdateCustomProperty(this SolidEdgeFileProperties.PropertySets propertySets, string propertyName, object propertyValue)
        {
            SolidEdgeFileProperties.Properties customPropertySet = null;
            SolidEdgeFileProperties.Property property = null;

            try
            {
                // Get a reference to the custom property set.
                customPropertySet = (SolidEdgeFileProperties.Properties)propertySets["Custom"];

                // Attempt to get the custom property.
                property = (SolidEdgeFileProperties.Property)customPropertySet[propertyName];

                // If we get here, the custom property exists so update the value.
                property.Value = propertyValue;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // If we get here, the custom property does not exist so add it.
                property = (SolidEdgeFileProperties.Property)customPropertySet.Add(propertyName, propertyValue);
            }
            catch
            {
                throw;
            }

            return property;
        }

        /// <summary>
        /// Adds or modifies a custom property.
        /// </summary>
        public static SolidEdgeFileProperties.Property UpdateCustomProperty(this SolidEdgeFileProperties.PropertySets propertySets, string propertyName, bool propertyValue)
        {
            return propertySets.InternalUpdateCustomProperty(propertyName, propertyValue);
        }

        /// <summary>
        /// Adds or modifies a custom property.
        /// </summary>
        public static SolidEdgeFileProperties.Property UpdateCustomProperty(this SolidEdgeFileProperties.PropertySets propertySets, string propertyName, DateTime propertyValue)
        {
            return propertySets.InternalUpdateCustomProperty(propertyName, propertyValue);
        }

        /// <summary>
        /// Adds or modifies a custom property.
        /// </summary>
        public static SolidEdgeFileProperties.Property UpdateCustomProperty(this SolidEdgeFileProperties.PropertySets propertySets, string propertyName, double propertyValue)
        {
            return propertySets.InternalUpdateCustomProperty(propertyName, propertyValue);
        }

        /// <summary>
        /// Adds or modifies a custom property.
        /// </summary>
        public static SolidEdgeFileProperties.Property UpdateCustomProperty(this SolidEdgeFileProperties.PropertySets propertySets, string propertyName, int propertyValue)
        {
            return propertySets.InternalUpdateCustomProperty(propertyName, propertyValue);
        }

        /// <summary>
        /// Adds or modifies a custom property.
        /// </summary>
        public static SolidEdgeFileProperties.Property UpdateCustomProperty(this SolidEdgeFileProperties.PropertySets propertySets, string propertyName, string propertyValue)
        {
            return propertySets.InternalUpdateCustomProperty(propertyName, propertyValue);
        }
    }
}
