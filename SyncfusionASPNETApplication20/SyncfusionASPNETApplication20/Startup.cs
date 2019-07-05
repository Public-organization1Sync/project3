using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(SyncfusionASPNETApplication3.Startup))]
namespace SyncfusionASPNETApplication3
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
