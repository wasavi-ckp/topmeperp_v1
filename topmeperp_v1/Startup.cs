using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(topmeperp.Startup))]
namespace topmeperp
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
