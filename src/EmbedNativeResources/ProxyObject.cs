using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EmbedNativeResources
{
    public class ProxyObject : MarshalByRefObject
    {
        public ProxyObject()
        {
        }

        public NativeResource[] GetNativeResourceInfo(string assemblyPath)
        {
            List<NativeResource> list = new List<NativeResource>();

            var assembly = Assembly.LoadFrom(assemblyPath);
            var attributes = assembly
                .GetCustomAttributes(false)
                .Where(x => x.GetType().FullName.Equals("SolidEdgeCommunity.AddIn.NativeResourceAttribute"));

            foreach (dynamic attribute in attributes)
            {
                NativeResource item = new NativeResource();

                item.Id = attribute.Id;
                item.Path = attribute.Path;
                list.Add(item);
            }

            return list.ToArray();
        }
    }
}
