using SolidEdgeCommunity;
using SolidEdgeCommunity.AddIn;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace TestAddIn
{
    [ComVisible(true)]
    [Guid("00000000-0000-0000-0000-C7823CD102E4")]
    [ProgId("SolidEdge.Community.TestAddIn.MyAddIn")]
    public class MyAddIn : SolidEdgeCommunity.AddIn.SolidEdgeAddIn
    {
        public override void OnConnection(SolidEdgeFramework.Application application, SolidEdgeFramework.SeConnectMode ConnectMode, SolidEdgeFramework.AddIn AddInInstance)
        {
            base.OnConnection(application, ConnectMode, AddInInstance);

            // If you makes changes to your ribbon, be sure to increment the GuiVersion. That makes bFirstTime = true
            // next time Solid Edge is started. bFirstTime is used to setup the ribbon so if you make a change but don't
            // see the changes, that could be why
            AddInEx.GuiVersion = 1;

            // Put your custom OnConnection code here.
            var applicationEvents = (SolidEdgeFramework.ISEApplicationEvents_Event)application.ApplicationEvents;
            applicationEvents.AfterWindowActivate += applicationEvents_AfterWindowActivate;
        }

        void applicationEvents_AfterWindowActivate(object theWindow)
        {
            var window = (SolidEdgeFramework.Window)theWindow;
            this.ViewOverlayController.Add<MyViewOverlay>(window.View);
        }

        public override void OnConnectToEnvironment(SolidEdgeFramework.Environment environment, bool firstTime)
        {
        }

        public override void OnDisconnection(SolidEdgeFramework.SeDisconnectMode DisconnectMode)
        {
        }

        public override void OnCreateRibbon(RibbonController controller, Guid environmentCategory, bool firstTime)
        {
            controller.Add<MyRibbon>(environmentCategory, firstTime);
        }

        public override void OnCreateEdgeBarPage(EdgeBarController controller, SolidEdgeFramework.SolidEdgeDocument document)
        {
            controller.Add<MyEdgeBarControl>(document, 1);
        }

        #region Registration functions

        [ComRegisterFunction]
        public static void OnRegister(Type t)
        {
            // See http://www.codeproject.com/Articles/839585/Solid-Edge-ST-AddIn-Architecture-Overview#Registration for registration details.
            // The following code helps write registry entries that Solid Edge needs to identify an addin. You can omit this code and
            // user installer logic if you'd like. This is simply here to help.

            try
            {
                var settings = new RegistrationSettings(t);

                settings.Enabled = true;
                settings.Environments.Add(SolidEdgeSDK.EnvironmentCategories.Application);
                settings.Environments.Add(SolidEdgeSDK.EnvironmentCategories.AllDocumentEnvrionments);

                // See http://msdn.microsoft.com/en-us/goglobal/bb964664.aspx for LCID details.
                var englishCulture = CultureInfo.GetCultureInfo(1033);

                // Title & Summary are Locale specific. 
                settings.Titles.Add(englishCulture, "SolidEdge.Community.TestAddIn");
                settings.Summaries.Add(englishCulture, "Solid Edge Addin in .NET 4.0.");

                var spanishCultere = CultureInfo.GetCultureInfo(3082);
                settings.Titles.Add(spanishCultere, "SolidEdge.Community.TestAddIn");
                settings.Summaries.Add(spanishCultere, "Solid Edge Addin in .NET 4.0.");

                // Optionally, you can add additional locales.
                var germanCultere = CultureInfo.GetCultureInfo(1031);
                settings.Titles.Add(germanCultere, "SolidEdge.Community.TestAddIn");
                settings.Summaries.Add(germanCultere, "Solid Edge Addin in .NET 4.0.");

                MyAddIn.Register(settings);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.StackTrace, ex.Message);
            }
        }

        [ComUnregisterFunction]
        public static void OnUnregister(Type t)
        {
            MyAddIn.Unregister(t);
        }

        #endregion
    }
}