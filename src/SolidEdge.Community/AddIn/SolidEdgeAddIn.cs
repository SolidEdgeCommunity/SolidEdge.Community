using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace SolidEdgeCommunity.AddIn
{
    public abstract class SolidEdgeAddIn : MarshalByRefObject, SolidEdgeFramework.ISolidEdgeAddIn
    {
        private static SolidEdgeAddIn _instance;
        private AppDomain _isolatedDomain;
        private SolidEdgeFramework.ISolidEdgeAddIn _isolatedAddIn;
        private SolidEdgeFramework.Application _application;
        private SolidEdgeFramework.AddIn _addInInstance;
        private RibbonController _ribbonController;
        private EdgeBarController _edgeBarController;
        private ViewOverlayController _viewOverlayController;

        /// <summary>
        /// Public constructor
        /// </summary>
        public SolidEdgeAddIn()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        #region SolidEdgeFramework.ISolidEdgeAddIn implementation

        void SolidEdgeFramework.ISolidEdgeAddIn.OnConnection(object Application, SolidEdgeFramework.SeConnectMode ConnectMode, SolidEdgeFramework.AddIn AddInInstance)
        {
            if (IsDefaultAppDomain)
            {
                _application = (SolidEdgeFramework.Application)Application;
                _addInInstance = AddInInstance;

                // Notice that "\n" is prepended to the description. This allows the addin to have its own Ribbon Tabs.
                this.AddInEx.Description = String.Format("\n{0}", this.AddInEx.Description);

                InitializeIsolatedAddIn();

                if (_isolatedAddIn != null)
                {
                    // Forward call to isolated AppDomain.
                    _isolatedAddIn.OnConnection(_application, ConnectMode, AddInInstance);
                }
            }
            else
            {
                Application = UnwrapTransparentProxy(Application);
                AddInInstance = UnwrapTransparentProxy<SolidEdgeFramework.AddIn>(AddInInstance);

                _instance = this;
                _application = (SolidEdgeFramework.Application)Application;
                _addInInstance = AddInInstance;
                _ribbonController = new RibbonController(this);
                _edgeBarController = new EdgeBarController(this);
                _viewOverlayController = new ViewOverlayController();

                OnConnection(_application, ConnectMode, AddInInstance);
            }
        }

        void SolidEdgeFramework.ISolidEdgeAddIn.OnConnectToEnvironment(string EnvCatID, object pEnvironmentDispatch, bool bFirstTime)
        {
            if ((IsDefaultAppDomain) && (_isolatedAddIn != null))
            {
                // Forward call to isolated AppDomain.
                _isolatedAddIn.OnConnectToEnvironment(EnvCatID, pEnvironmentDispatch, bFirstTime);
            }
            else
            {
                pEnvironmentDispatch = UnwrapTransparentProxy(pEnvironmentDispatch);
                var environment = (SolidEdgeFramework.Environment)pEnvironmentDispatch;
                var environmentCategory = new Guid(EnvCatID);

                OnConnectToEnvironment(environment, bFirstTime);
                OnCreateRibbon(_ribbonController, environmentCategory, bFirstTime);
            }
        }

        void SolidEdgeFramework.ISolidEdgeAddIn.OnDisconnection(SolidEdgeFramework.SeDisconnectMode DisconnectMode)
        {
            if (IsDefaultAppDomain)
            {
                if (_isolatedAddIn != null)
                {
                    // Forward call to isolated AppDomain.
                    _isolatedAddIn.OnDisconnection(DisconnectMode);

                    // Unload isolated domain.
                    if (_isolatedDomain != null)
                    {
                        AppDomain.Unload(_isolatedDomain);
                    }

                    _isolatedDomain = null;
                    _isolatedAddIn = null;
                }
            }
            else
            {
                OnDisconnection(DisconnectMode);

                if (_ribbonController != null)
                {
                    try
                    {
                        _ribbonController.Dispose();
                    }
                    catch
                    {
                    }

                    _ribbonController = null;
                }

                if (_edgeBarController != null)
                {
                    try
                    {
                        _edgeBarController.Dispose();
                    }
                    catch
                    {
                    }

                    _edgeBarController = null;
                }

                if (_viewOverlayController != null)
                {
                    try
                    {
                        _viewOverlayController.Dispose();
                    }
                    catch
                    {
                    }

                    _viewOverlayController = null;
                }

            }
        }

        #endregion

        #region SolidEdge.Community.AddIn.SolidEdgeAddIn virtual members

        /// <summary>
        /// 
        /// </summary>
        public virtual void OnCreateRibbon(RibbonController controller, Guid environmentCategory, bool firstTime)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="controller"></param>
        /// <param name="document"></param>
        public virtual void OnCreateEdgeBarPage(EdgeBarController controller, SolidEdgeFramework.SolidEdgeDocument document)
        {
        }

        /// <summary>
        /// Called by SolidEdgeFramework.ISolidEdgeAddIn.OnConnection().
        /// </summary>
        /// <param name="application"></param>
        /// <param name="ConnectMode"></param>
        /// <param name="AddInInstance"></param>
        public virtual void OnConnection(SolidEdgeFramework.Application application, SolidEdgeFramework.SeConnectMode ConnectMode, SolidEdgeFramework.AddIn AddInInstance)
        {
        }

        /// <summary>
        /// Called by SolidEdgeFramework.ISolidEdgeAddIn.OnConnectToEnvironment().
        /// </summary>
        public virtual void OnConnectToEnvironment(SolidEdgeFramework.Environment environment, bool firstTime)
        {
        }

        /// <summary>
        /// Called by SolidEdgeFramework.ISolidEdgeAddIn.OnDisconnection().
        /// </summary>
        public virtual void OnDisconnection(SolidEdgeFramework.SeDisconnectMode DisconnectMode)
        {
        }

        /// <summary>
        /// Returns the path of the .dll\.exe containing native Win32 resources. Typically the current assembly location.
        /// </summary>
        /// <remarks>
        /// It is only necessary to override if you have native resources in a separate .dll.
        /// </remarks>
        public virtual string NativeResourcesDllPath
        {
            get { return this.GetType().Assembly.Location; }
        }

        #endregion

        #region Methods

        private void InitializeIsolatedAddIn()
        {
            var type = this.GetType();

            AppDomainSetup appDomainSetup = AppDomain.CurrentDomain.SetupInformation;
            Evidence evidence = AppDomain.CurrentDomain.Evidence;
            appDomainSetup.ApplicationBase = Path.GetDirectoryName(type.Assembly.Location);

            string domainName = this.AddIn.GUID;
            _isolatedDomain = AppDomain.CreateDomain(domainName, evidence, appDomainSetup);
            var proxyAddIn = _isolatedDomain.CreateInstanceAndUnwrap(type.Assembly.FullName, type.FullName);
            _isolatedAddIn = (SolidEdgeFramework.ISolidEdgeAddIn)proxyAddIn;
        }

        /// <summary>
        /// Implementation of MarshalByRefObject.InitializeLifetimeService(). Not intended to be called directly.
        /// </summary>
        /// <returns></returns>
        public override object InitializeLifetimeService()
        {
            return null;
        }

        Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            var assemblies = AppDomain.CurrentDomain.GetAssemblies();

            foreach (Assembly assembly in assemblies)
            {
                var assemblyName = assembly.GetName();
                var thatAssemblyName = new AssemblyName(args.Name);

                if (assembly.FullName.Equals(args.Name))
                {
                    return assembly;
                }
            }

            return null;
        }

        private object UnwrapTransparentProxy(object rcw)
        {
            if (RemotingServices.IsTransparentProxy(rcw))
            {
                IntPtr punk = Marshal.GetIUnknownForObject(rcw);

                try
                {
                    return Marshal.GetObjectForIUnknown(punk);
                }
                finally
                {
                    Marshal.Release(punk);
                }
            }

            return rcw;
        }

        private TInterface UnwrapTransparentProxy<TInterface>(object rcw) where TInterface : class
        {
            if (RemotingServices.IsTransparentProxy(rcw))
            {
                IntPtr punk = Marshal.GetIUnknownForObject(rcw);

                try
                {
                    return (TInterface)Marshal.GetObjectForIUnknown(punk);
                }
                finally
                {
                    Marshal.Release(punk);
                }
            }

            return rcw as TInterface;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.Application.
        /// </summary>
        public SolidEdgeFramework.Application Application { get { return _application; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISEApplicationEvents_Event.
        /// </summary>
        public SolidEdgeFramework.ISEApplicationEvents_Event ApplicationEvents { get { return _application.ApplicationEvents as SolidEdgeFramework.ISEApplicationEvents_Event; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISEApplicationWindowEvents_Event.
        /// </summary>
        public SolidEdgeFramework.ISEApplicationWindowEvents_Event ApplicationWindowEvents { get { return _application.ApplicationWindowEvents as SolidEdgeFramework.ISEApplicationWindowEvents_Event; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISEAddIn.
        /// </summary>
        public SolidEdgeFramework.ISEAddIn AddIn { get { return _addInInstance as SolidEdgeFramework.ISEAddIn; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISEAddInEvents_Event;
        /// </summary>
        public SolidEdgeFramework.ISEAddInEvents_Event AddInEvents { get { return _addInInstance.AddInEvents as SolidEdgeFramework.ISEAddInEvents_Event; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISEAddInEx.
        /// </summary>
        /// <remarks>
        /// ISEAddInEx is available in Solid Edge ST or newer.
        /// </remarks>
        public SolidEdgeFramework.ISEAddInEx AddInEx { get { return _addInInstance as SolidEdgeFramework.ISEAddInEx; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISEAddIn.
        /// </summary>
        /// <remarks>
        /// ISEAddInEx is available in Solid Edge ST7 or newer.
        /// </remarks>
        public SolidEdgeFramework.ISEAddInEx2 AddInEx2 { get { return _addInInstance as SolidEdgeFramework.ISEAddInEx2; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISolidEdgeBar.
        /// </summary>
        public SolidEdgeFramework.ISolidEdgeBar EdgeBar { get { return _addInInstance as SolidEdgeFramework.ISolidEdgeBar; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISolidEdgeBarEx.
        /// </summary>
        public SolidEdgeFramework.ISolidEdgeBarEx EdgeBarEx { get { return _addInInstance as SolidEdgeFramework.ISolidEdgeBarEx; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISolidEdgeBarEx2.
        /// </summary>
        public SolidEdgeFramework.ISolidEdgeBarEx2 EdgeBarEx2 { get { return _addInInstance as SolidEdgeFramework.ISolidEdgeBarEx2; } }

        /// <summary>
        /// Returns an instance of EdgeBarController.
        /// </summary>
        public EdgeBarController EdgeBarController { get { return _edgeBarController; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISEFeatureLibraryEvents_Event.
        /// </summary>
        public SolidEdgeFramework.ISEFeatureLibraryEvents_Event FeatureLibraryEvents { get { return _application.FeatureLibraryEvents as SolidEdgeFramework.ISEFeatureLibraryEvents_Event; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISEFileUIEvents_Event.
        /// </summary>
        public SolidEdgeFramework.ISEFileUIEvents_Event FileUIEvents { get { return _application.FileUIEvents as SolidEdgeFramework.ISEFileUIEvents_Event; } }

        /// <summary>
        /// Returns the GUID of the AddIn.
        /// </summary>
        public Guid Guid { get { return new Guid(_addInInstance.GUID); } }

        /// <summary>
        /// Returns the version of the GUI for the AddIn.
        /// </summary>
        public int GuiVersion { get { return _addInInstance.GuiVersion; } }

        /// <summary>
        /// Returns a global static instance of the addin.
        /// </summary>
        public static SolidEdgeAddIn Instance { get { return _instance; } }

        bool IsDefaultAppDomain { get { return AppDomain.CurrentDomain.IsDefaultAppDomain(); } }
        bool IsIsolatedAppDomain { get { return !AppDomain.CurrentDomain.IsDefaultAppDomain(); } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISENewFileUIEvents_Event.
        /// </summary>
        public SolidEdgeFramework.ISENewFileUIEvents_Event NewFileUIEvents { get { return _application.NewFileUIEvents as SolidEdgeFramework.ISENewFileUIEvents_Event; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISEECEvents_Event.
        /// </summary>
        public SolidEdgeFramework.ISEECEvents_Event SEECEvents { get { return _application.SEECEvents as SolidEdgeFramework.ISEECEvents_Event; } }

        /// <summary>
        /// Returns an instance of SolidEdgeFramework.ISEShortCutMenuEvents_Event.
        /// </summary>
        public SolidEdgeFramework.ISEShortCutMenuEvents_Event ShortcutMenuEvents { get { return _application.ShortcutMenuEvents as SolidEdgeFramework.ISEShortCutMenuEvents_Event; } }

        /// <summary>
        /// Returns an instance of the ViewOverlayController.
        /// </summary>
        public ViewOverlayController ViewOverlayController { get { return _viewOverlayController; } }

        /// <summary>
        /// Returns an instance of RibbonController.
        /// </summary>
        public RibbonController RibbonController { get { return _ribbonController; } }

        #endregion

        #region regasm.exe functions

        /// <summary>
        /// Registers the addin.
        /// </summary>
        protected static void Register(Type t, string title, string summary, Guid[] environments)
        {
            Register(t, title, summary, true, environments);
        }

        /// <summary>
        /// Registers the addin.
        /// </summary>
        protected static void Register(Type t, string title, string summary, bool enabled, Guid[] environments)
        {
            var assembly = t.Assembly;

            #region HKEY_CLASSES_ROOT\CLSID\{GUID}

            RegisterImplementedCategories(t);
            RegisterOptions(t, enabled);

            #endregion

            #region HKEY_CLASSES_ROOT\CLSID\{GUID}\Environment Categories

            //List<Guid> catids = new List<Guid>();

            //foreach (var environment in manifest.Registration.Environments)
            //{
            //    var catid = environment.CategoryIdAsGuid;
            //    if (catid.Equals(Guid.Empty) == false)
            //    {
            //        catids.Add(catid);
            //    }
            //}

            RegisterEnvironments(t, environments);

            #endregion

            #region HKEY_CLASSES_ROOT\CLSID\{GUID}\Summary

            var culture = CultureInfo.CurrentCulture;

            RegisterTitle(t, culture, title);
            RegisterSummary(t, culture, summary);

            #endregion
        }

        /// <summary>
        /// Deletes the HKEY_CLASSES_ROOT\CLSID\{ADDIN_GUID} key.
        /// </summary>
        /// <param name="t"></param>
        protected static void Unregister(Type t)
        {
            string subkey = String.Format(@"CLSID\{0}", t.GUID.ToRegistryString());
            Registry.ClassesRoot.DeleteSubKeyTree(subkey, false);
        }

        /// <summary>
        /// Creates or opens base registry key for addin.
        /// </summary>
        /// <param name="guid"></param>
        /// <returns></returns>
        static RegistryKey CreateBaseKey(Guid guid)
        {
            string subkey = String.Format(@"CLSID\{0}", guid.ToRegistryString());
            return Registry.ClassesRoot.CreateSubKey(subkey);
        }

        /// <summary>
        /// Registers the environments that the addin is subscribed to.
        /// </summary>
        /// <param name="t"></param>
        /// <param name="environments"></param>
        static void RegisterEnvironments(Type t, Guid[] environments)
        {
            using (RegistryKey baseKey = CreateBaseKey(t.GUID))
            {
                foreach (var environment in environments)
                {
                    var subkey = String.Format(@"Environment Categories\{0}", environment.ToRegistryString());

                    using (RegistryKey environmentCategoryKey = baseKey.CreateSubKey(subkey))
                    {
                    }
                }
            }
        }

        /// <summary>
        /// Registers the implemented categories of the addin.
        /// </summary>
        /// <param name="t"></param>
        static void RegisterImplementedCategories(Type t)
        {
            using (RegistryKey baseKey = CreateBaseKey(t.GUID))
            {
                var guid = new Guid(SolidEdgeSDK.CATID.SolidEdgeAddIn);
                var subkey = String.Format(@"Implemented Categories\{0}", guid.ToRegistryString());

                using (RegistryKey implementedCategoriesKey = baseKey.CreateSubKey(subkey))
                {
                }
            }
        }

        /// <summary>
        /// Registers the addin options.
        /// </summary>
        /// <param name="t"></param>
        /// <param name="autoConnect"></param>
        static void RegisterOptions(Type t, bool autoConnect)
        {
            using (RegistryKey baseKey = CreateBaseKey(t.GUID))
            {
                baseKey.SetValue("AutoConnect", autoConnect ? 1 : 0);
            }
        }

        static void RegisterTitle(Type t, CultureInfo culture, string title)
        {
            // Example Local ID (LCID)
            // Description: English - United States
            // int: 1033
            // hex: 0x0409
            // HKEY_CLASSES_ROOT\CLSID\{ADDIN_GUID}\409

            int hexLCID = int.Parse(culture.LCID.ToString("X4"));
            string keyName = hexLCID.ToString();

            using (RegistryKey baseKey = CreateBaseKey(t.GUID))
            {
                // Write the title value.
                baseKey.SetValue(keyName, title);
            }
        }

        static void RegisterSummary(Type t, CultureInfo culture, string summary)
        {
            // Example Local ID (LCID)
            // Description: English - United States
            // int: 1033
            // hex: 0x0409
            // HKEY_CLASSES_ROOT\CLSID\{ADDIN_GUID}\Summary\409

            int hexLCID = int.Parse(culture.LCID.ToString("X4"));
            string keyName = hexLCID.ToString();

            using (RegistryKey baseKey = CreateBaseKey(t.GUID))
            {
                // Write the summary key.
                using (RegistryKey summaryKey = baseKey.CreateSubKey("Summary"))
                {
                    summaryKey.SetValue(keyName, summary);
                }
            }
        }

        #endregion
    }
}
