using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Optimization;
using System.Web.Routing;
using System.Web.Security;
using System.Web.SessionState;
using Syncfusion.Licensing;

namespace SyncfusionASPNETApplication27
{
    public class Global : HttpApplication
    {
        void Application_Start(object sender, EventArgs e)
        {
			//Syncfusion Licensing Register
	        SyncfusionLicenseProvider.RegisterLicense("MDAxQDMxMzcyZTMyMmUzMFVWUDA4MDl0UlBNNnZSdndOa2xPaUI3Z09QSU1GbGQvZ1dyYVlubUY3b1E9");
            // Code that runs on application startup
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
        }
    }
}