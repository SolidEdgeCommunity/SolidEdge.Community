using SolidEdgeCommunity.Runtime.InteropServices.ComTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace SolidEdgeCommunity.Runtime.InteropServices
{
    /// <summary>
    /// COM object wrapper class.
    /// </summary>
    public class ComObject
    {
        const int LOCALE_SYSTEM_DEFAULT = 2048;

        /// <summary>
        /// Using IDispatch, returns the ITypeInfo of the specified object.
        /// </summary>
        /// <param name="comObject"></param>
        /// <returns></returns>
        public static ITypeInfo GetITypeInfo(object comObject)
        {
            if (Marshal.IsComObject(comObject) == false) throw new InvalidComObjectException();

            var dispatch = comObject as IDispatch;

            if (dispatch != null)
            {
                return dispatch.GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT);
            }

            return null;
        }

        /// <summary>
        /// Returns a strongly typed property by name using the specified COM object.
        /// </summary>
        /// <typeparam name="T">The type of the property to return.</typeparam>
        /// <param name="comObject"></param>
        /// <param name="name">The name of the property to retrieve.</param>
        /// <returns></returns>
        public static T GetPropertyValue<T>(object comObject, string name)
        {
            if (Marshal.IsComObject(comObject) == false) throw new InvalidComObjectException();

            var type = comObject.GetType();
            var value = type.InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, comObject, null);

            return (T)value;
        }

        /// <summary>
        /// Returns a strongly typed property by name using the specified COM object.
        /// </summary>
        /// <typeparam name="T">The type of the property to return.</typeparam>
        /// <param name="comObject"></param>
        /// <param name="name">The name of the property to retrieve.</param>
        /// <param name="defaultValue">The value to return if the property does not exist.</param>
        /// <returns></returns>
        public static T GetPropertyValue<T>(object comObject, string name, T defaultValue)
        {
            if (Marshal.IsComObject(comObject) == false) throw new InvalidComObjectException();

            var type = comObject.GetType();

            try
            {
                var value = type.InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, comObject, null);
                return (T)value;
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Using IDispatch, determine the managed type of the specified object.
        /// </summary>
        /// <param name="comObject"></param>
        /// <returns></returns>
        public static Type GetType(object comObject)
        {
            if (Marshal.IsComObject(comObject) == false) throw new InvalidComObjectException();

            Type type = null;
            var dispatch = comObject as IDispatch;
            ITypeInfo typeInfo = null;
            var pTypeAttr = IntPtr.Zero;
            var typeAttr = default(System.Runtime.InteropServices.ComTypes.TYPEATTR);

            try
            {
                if (dispatch != null)
                {
                    typeInfo = dispatch.GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT);
                    typeInfo.GetTypeAttr(out pTypeAttr);
                    typeAttr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)Marshal.PtrToStructure(pTypeAttr, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));

                    // Type can technically be defined in any loaded assembly.
                    var assemblies = AppDomain.CurrentDomain.GetAssemblies();

                    // Scan each assembly for a type with a matching GUID.
                    foreach (var assembly in assemblies)
                    {
                        type = assembly.GetTypes()
                            .Where(x => x.IsInterface)
                            .Where(x => x.GUID.Equals(typeAttr.guid))
                            .FirstOrDefault();

                        if (type != null)
                        {
                            // Found what we're looking for so break out of the loop.
                            break;
                        }
                    }
                }
            }
            finally
            {
                if (typeInfo != null)
                {
                    typeInfo.ReleaseTypeAttr(pTypeAttr);
                    Marshal.ReleaseComObject(typeInfo);
                }
            }

            return type;
        }
    }
}
