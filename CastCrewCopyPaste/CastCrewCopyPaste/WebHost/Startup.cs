using System.Web.Http;
using Owin;

namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste.WebHost
{
    internal sealed class Startup
    {
        internal const string HostBinding = "http://localhost:10001";

        public void Configuration(IAppBuilder app)
        {
            var config = new HttpConfiguration();

            SwaggerConfig.RegisterOwin(config);

            WebApiConfig.Register(config);

            app.UseWebApi(config);
        }
    }
}
