using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Optimization;
using System.Web.Routing;
using System.Web.Security;
using System.Web.SessionState;
using Syncfusion.Licensing;

namespace SyncfusionASPNETApplication3
{
    public class Global : HttpApplication
    {
        void Application_Start(object sender, EventArgs e)
        {
	    //Syncfusion Licensing Register
	    SyncfusionLicenseProvider.RegisterLicense("MDAxQDMxMzcyZTMyMmUzMFpKNkVURjVrVzZoWTZKbTBTSEFONE9ES2d0Q0dDc0UweUNyRVhtNU9ZS0E9");
            // Code that runs on application startup
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            System.Web.Http.GlobalConfiguration.Configuration.Routes.MapHttpRoute(
            name: "DefaultApi",
            routeTemplate: "api/{controller}/{action}/{id}",
            defaults: new { id = RouteParameter.Optional });
        }
    }
}