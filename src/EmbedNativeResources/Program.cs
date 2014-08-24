using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;

namespace EmbedNativeResources
{
    class Program
    {
        [DllImport("kernel32.dll", EntryPoint = "BeginUpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        static extern IntPtr BeginUpdateResource(string pFileName, bool bDeleteExistingResources);

        [DllImport("kernel32.dll", EntryPoint = "UpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        static extern bool UpdateResource(IntPtr hUpdate, IntPtr lpType, IntPtr lpName, UInt16 wLanguage, byte[] lpData, UInt32 cbData);

        [DllImport("kernel32.dll", EntryPoint = "EndUpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        static extern bool EndUpdateResource(IntPtr hUpdate, bool fDiscard);

        static int Main(string[] args)
        {
            try
            {
                if (args.Length == 2)
                {
                    var projectDir = args[0];
                    var assemblyPath = args[1];
                    //var addinManifest = Path.Combine(projectDir, "addin.manifest");

                    // Make sure specified project directory exists.
                    if (Directory.Exists(projectDir) == false)
                    {
                        throw new DirectoryNotFoundException(String.Format("'{0}' does not exist.", projectDir));
                    }

                    // Make sure specified assembly file exists.
                    if (File.Exists(assemblyPath) == false)
                    {
                        throw new FileNotFoundException(String.Format("'{0}' does not exist.", assemblyPath));
                    }

                    UpdateAssembly(projectDir, assemblyPath);

                    //if (File.Exists(addinManifest))
                    //{
                    //    UpdateAssembly(addinManifest, projectDir, assemblyPath);
                    //}

                    return 0;
                }
                else
                {
                    ShowUsage();
                    //return 2;
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
                //return 1;
            }

            return 0;
        }

        static void UpdateAssembly(string projectPath, string assemblyPath)
        {
            AppDomain isolatedAppDomain = null;
            NativeResource[] nativeResources = { };

            try
            {
                isolatedAppDomain = AppDomain.CreateDomain("IsolatedAppDomain");
                Type proxyType = typeof(ProxyObject);
                var proxyObject = isolatedAppDomain.CreateInstanceAndUnwrap(
                    proxyType.Assembly.FullName,
                    proxyType.FullName) as ProxyObject;

                nativeResources = proxyObject.GetNativeResourceInfo(assemblyPath);
            }
            catch
            {
                throw;
            }
            finally
            {
                if (isolatedAppDomain != null)
                {
                    AppDomain.Unload(isolatedAppDomain);
                }
            }

            if (nativeResources.Length > 0)
            {
                // Being the update process.
                IntPtr pAssembly = BeginUpdateResource(assemblyPath, false);

                if (pAssembly == IntPtr.Zero)
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }

                foreach (var nativeResource in nativeResources)
                {
                    var resourceId = nativeResource.Id;
                    var resourcePath = Path.Combine(projectPath, nativeResource.Path);
                    var resourceExtension = Path.GetExtension(resourcePath).ToUpper();
                    IntPtr pType = IntPtr.Zero;

                    if (resourceExtension == null) continue;

                    if (File.Exists(resourcePath))
                    {
                        var data = File.ReadAllBytes(resourcePath);

                        if (resourceExtension.Equals(".BMP", StringComparison.OrdinalIgnoreCase))
                        {
                            // Strip BITMAPFILEHEADER from bytes.
                            data = GetDibBytes(data);
                            pType = new IntPtr((int)ResourceType.RT_BITMAP);
                        }
                        else if (resourceExtension.Equals(".PNG", StringComparison.OrdinalIgnoreCase))
                        {
                            pType = Marshal.StringToHGlobalUni("PNG");
                        }

                        IntPtr pName = new IntPtr(resourceId);

                        if (!UpdateResource(pAssembly, pType, pName, 0, data, (data == null ? 0 : (uint)data.Length)))
                        {
                            throw new Win32Exception(Marshal.GetLastWin32Error());
                        }
                    }
                    else
                    {
                        Console.WriteLine("'{0}' does not exist.", resourcePath);
                    }
                }

                if (!EndUpdateResource(pAssembly, false))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
            }
        }

        #region Old code

        //static void UpdateAssembly(string manifestPath, string projectPath, string assemblyPath)
        //{
        //    XDocument xml = XDocument.Load(manifestPath);

        //    var resourceFolders = xml.Root.Descendants("Resources");

        //    if (resourceFolders.Count() > 0)
        //    {
        //        // Being the update process.
        //        IntPtr pAssembly = BeginUpdateResource(assemblyPath, false);

        //        if (pAssembly == IntPtr.Zero)
        //            throw new Win32Exception(Marshal.GetLastWin32Error());

        //        foreach (var resourceFolder in resourceFolders)
        //        {
        //            var pathAttribute = resourceFolder.Attribute("path");

        //            if (pathAttribute != null)
        //            {
        //                var resourcesPath = Path.Combine(projectPath, pathAttribute.Value);

        //                List<XElement> resources = new List<XElement>();
        //                resources.AddRange(resourceFolder.Descendants("PNG"));
        //                resources.AddRange(resourceFolder.Descendants("BMP"));

        //                foreach (var resource in resources)
        //                {
        //                    IntPtr pType = IntPtr.Zero;
        //                    var idAttribute = resource.Attribute("id");
        //                    var filenameAttribute = resource.Attribute("filename");

        //                    if ((idAttribute != null) && (filenameAttribute != null))
        //                    {
        //                        int id = 0;

        //                        if (int.TryParse(idAttribute.Value, out id))
        //                        {
        //                            var nativeResourceFile = new FileInfo(Path.Combine(resourcesPath, filenameAttribute.Value));

        //                            // Ensure that file exists.
        //                            if (nativeResourceFile.Exists)
        //                            {
        //                                var data = File.ReadAllBytes(nativeResourceFile.FullName);

        //                                if (resource.Name.ToString().Equals("BMP", StringComparison.OrdinalIgnoreCase))
        //                                {
        //                                    // Strip BITMAPFILEHEADER from bytes.
        //                                    data = GetDibBytes(data);
        //                                    pType = new IntPtr((int)ResourceType.RT_BITMAP);
        //                                }
        //                                else if (resource.Name.ToString().Equals("PNG", StringComparison.OrdinalIgnoreCase))
        //                                {
        //                                    pType = Marshal.StringToHGlobalUni("PNG");
        //                                }

        //                                IntPtr pName = new IntPtr(id);

        //                                if (!UpdateResource(pAssembly, pType, pName, 0, data, (data == null ? 0 : (uint)data.Length)))
        //                                {
        //                                    throw new Win32Exception(Marshal.GetLastWin32Error());
        //                                }
        //                            }
        //                            else
        //                            {
        //                                Console.WriteLine("'{0}' does not exist.", nativeResourceFile);
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }

        //        if (!EndUpdateResource(pAssembly, false))
        //            throw new Win32Exception(Marshal.GetLastWin32Error());
        //    }
        //}

        #endregion

        static void ShowUsage()
        {
            Console.WriteLine("EmbedNativeResources Version: {0}", Assembly.GetExecutingAssembly().GetName().Version);
            Console.WriteLine("usage: EmbedNativeResources.exe <project directory> <assembly full path>");
        }

        static byte[] GetDibBytes(byte[] bitmap)
        {
            Int32 size = bitmap.Length - Marshal.SizeOf(typeof(BITMAPFILEHEADER));
            byte[] bitmapData = new byte[size];
            Buffer.BlockCopy(bitmap, Marshal.SizeOf(typeof(BITMAPFILEHEADER)), bitmapData, 0, size);
            return bitmapData;
        }
    }
}
