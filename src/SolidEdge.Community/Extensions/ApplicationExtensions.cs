using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SolidEdgeCommunity.Extensions
{
    /// <summary>
    /// SolidEdgeFramework.Application extension methods.
    /// </summary>
    public static class ApplicationExtensions
    {
        /// <summary>
        /// Creates a Solid Edge command control.
        /// </summary>
        /// <param name="application"></param>
        /// <param name="CmdFlags"></param>
        public static SolidEdgeFramework.Command CreateCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.seCmdFlag CmdFlags)
        {
            return application.CreateCommand((int)CmdFlags);
        }

        /// <summary>
        /// Returns the currently active document.
        /// </summary>
        /// <remarks>An exception will be thrown if there is no active document.</remarks>
        public static SolidEdgeFramework.SolidEdgeDocument GetActiveDocument(this SolidEdgeFramework.Application application)
        {
            return application.GetActiveDocument(true);
        }

        /// <summary>
        /// Returns the currently active document.
        /// </summary>
        public static SolidEdgeFramework.SolidEdgeDocument GetActiveDocument(this SolidEdgeFramework.Application application, bool throwOnError)
        {
            try
            {
                // ActiveDocument will throw an exception if no document is open.
                return (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
            }
            catch
            {
                if (throwOnError) throw;
            }

            return null;
        }

        /// <summary>
        /// Returns the currently active document.
        /// </summary>
        /// <typeparam name="T">The type to return.</typeparam>
        /// /// <remarks>An exception will be thrown if there is no active document or if the cast fails.</remarks>
        public static T GetActiveDocument<T>(this SolidEdgeFramework.Application application) where T : class
        {
            return application.GetActiveDocument<T>(true);
        }

        /// <summary>
        /// Returns the currently active document.
        /// </summary>
        /// <typeparam name="T">The type to return.</typeparam>
        public static T GetActiveDocument<T>(this SolidEdgeFramework.Application application, bool throwOnError) where T : class
        {
            try
            {
                // ActiveDocument will throw an exception if no document is open.
                return (T)application.ActiveDocument;
            }
            catch
            {
                if (throwOnError) throw;
            }

            return null;
        }

        /// <summary>
        /// Returns the environment that belongs to the current active window of the document.
        /// </summary>
        public static SolidEdgeFramework.Environment GetActiveEnvironment(this SolidEdgeFramework.Application application)
        {
            SolidEdgeFramework.Environments environments = application.Environments;
            return environments.Item(application.ActiveEnvironment);
        }

        ///// <summary>
        ///// Returns the application events.
        ///// </summary>
        //public static SolidEdgeFramework.ISEApplicationEvents_Event GetApplicationEvents(this SolidEdgeFramework.Application application)
        //{
        //    return (SolidEdgeFramework.ISEApplicationEvents_Event)application.ApplicationEvents;
        //}

        ///// <summary>
        ///// Returns the application window events.
        ///// </summary>
        //public static SolidEdgeFramework.ISEApplicationWindowEvents_Event GetApplicationWindowEvents(this SolidEdgeFramework.Application application)
        //{
        //    return (SolidEdgeFramework.ISEApplicationWindowEvents_Event)application.ApplicationWindowEvents;
        //}

        /// <summary>
        /// Returns an environment specified by CATID.
        /// </summary>
        public static SolidEdgeFramework.Environment GetEnvironment(this SolidEdgeFramework.Application application, string CATID)
        {
            var guid1 = new Guid(CATID);

            foreach (var environment in application.Environments.OfType<SolidEdgeFramework.Environment>())
            {
                var guid2 = new Guid(environment.CATID);
                if (guid1.Equals(guid2))
                {
                    return environment;
                }
            }

            return null;
        }

        ///// <summary>
        ///// Returns the feature library events.
        ///// </summary>
        //public static SolidEdgeFramework.ISEFeatureLibraryEvents_Event GetFeatureLibraryEvents(this SolidEdgeFramework.Application application)
        //{
        //    return (SolidEdgeFramework.ISEFeatureLibraryEvents_Event)application.FeatureLibraryEvents;
        //}

        ///// <summary>
        ///// Returns the file UI events.
        ///// </summary>
        //public static SolidEdgeFramework.ISEFileUIEvents_Event GetFileUIEvents(this SolidEdgeFramework.Application application)
        //{
        //    return (SolidEdgeFramework.ISEFileUIEvents_Event)application.FileUIEvents;
        //}

        /// <summary>
        /// Returns the value of a specified global constant.
        /// </summary>
        public static object GetGlobalParameter(this SolidEdgeFramework.Application application, SolidEdgeFramework.ApplicationGlobalConstants globalConstant)
        {
            object value = null;
            application.GetGlobalParameter(globalConstant, ref value);
            return value;
        }

        /// <summary>
        /// Returns the value of a specified global constant.
        /// </summary>
        /// <typeparam name="T">The type to return.</typeparam>
        public static T GetGlobalParameter<T>(this SolidEdgeFramework.Application application, SolidEdgeFramework.ApplicationGlobalConstants globalConstant)
        {
            object value = null;
            application.GetGlobalParameter(globalConstant, ref value);
            return (T)value;
        }

        /// <summary>
        /// Returns a NativeWindow object that represents the main application window.
        /// </summary>
        public static NativeWindow GetNativeWindow(this SolidEdgeFramework.Application application)
        {
            return NativeWindow.FromHandle(new IntPtr(application.hWnd));
        }

        ///// <summary>
        ///// Returns the new file UI events.
        ///// </summary>
        //public static SolidEdgeFramework.ISENewFileUIEvents_Event GetNewFileUIEvents(this SolidEdgeFramework.Application application)
        //{
        //    return (SolidEdgeFramework.ISENewFileUIEvents_Event)application.NewFileUIEvents;
        //}

        ///// <summary>
        ///// Returns the SEEC events.
        ///// </summary>
        //public static SolidEdgeFramework.ISEECEvents_Event GetSEECEvents(this SolidEdgeFramework.Application application)
        //{
        //    return (SolidEdgeFramework.ISEECEvents_Event)application.SEECEvents;
        //}

        /// <summary>
        /// Returns a Process object that represents the application prcoess.
        /// </summary>
        public static Process GetProcess(this SolidEdgeFramework.Application application)
        {
            return Process.GetProcessById(application.ProcessID);
        }

        ///// <summary>
        ///// Returns the shortcut menu events.
        ///// </summary>
        //public static SolidEdgeFramework.ISEShortCutMenuEvents_Event GetShortcutMenuEvents(this SolidEdgeFramework.Application application)
        //{
        //    return (SolidEdgeFramework.ISEShortCutMenuEvents_Event)application.ShortcutMenuEvents;
        //}

        /// <summary>
        /// Returns a Version object that represents the application version.
        /// </summary>
        public static Version GetVersion(this SolidEdgeFramework.Application application)
        {
            return new Version(application.Version);
        }

        /// <summary>
        /// Returns a point object to the application main window.
        /// </summary>
        public static IntPtr GetWindowHandle(this SolidEdgeFramework.Application application)
        {
            return new IntPtr(application.hWnd);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.AssemblyCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.CuttingPlaneLineCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.DetailCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.DrawingViewEditCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.ExplodeCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.LayoutCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.LayoutInPartCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.MotionCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.PartCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.ProfileCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.ProfileHoleCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.ProfilePatternCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.ProfileRevolvedCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.SheetMetalCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.SimplifyCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.SolidEdgeCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.StudioCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.TubingCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeConstants.WeldmentCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Activates a specified Solid Edge command.
        /// </summary>
        public static void StartCommand(this SolidEdgeFramework.Application application, SolidEdgeFramework.SolidEdgeCommandConstants CommandID)
        {
            application.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)CommandID);
        }

        /// <summary>
        /// Shows the form with the application main window as the owner.
        /// </summary>
        public static void Show(this SolidEdgeFramework.Application application, System.Windows.Forms.Form form)
        {
            if (form == null) throw new ArgumentNullException("form");

            form.Show(application.GetNativeWindow());
        }

        /// <summary>
        /// Shows the form as a modal dialog box with the application main window as the owner.
        /// </summary>
        public static DialogResult ShowDialog(this SolidEdgeFramework.Application application, System.Windows.Forms.Form form)
        {
            if (form == null) throw new ArgumentNullException("form");

            return form.ShowDialog(application.GetNativeWindow());
        }

        /// <summary>
        /// Shows the form as a modal dialog box with the application main window as the owner.
        /// </summary>
        public static DialogResult ShowDialog(this SolidEdgeFramework.Application application, CommonDialog dialog)
        {
            if (dialog == null) throw new ArgumentNullException("dialog");
            return dialog.ShowDialog(application.GetNativeWindow());
        }
    }
}
