using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(DbToXLSX.Startup))]
namespace DbToXLSX
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
