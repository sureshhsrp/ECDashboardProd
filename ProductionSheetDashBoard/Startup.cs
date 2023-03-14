using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ProductionSheetDashBoard.Startup))]
namespace ProductionSheetDashBoard
{
    public partial class Startup {
        public void Configuration(IAppBuilder app) {
            ConfigureAuth(app);
        }
    }
}
