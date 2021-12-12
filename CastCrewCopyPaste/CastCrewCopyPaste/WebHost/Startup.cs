﻿namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste.WebHost
{
    using System.Web.Http;
    using Owin;
    using Swashbuckle.Application;

    internal sealed class Startup
    {
        internal const string HostBinding = "http://localhost:10001";

        public void Configuration(IAppBuilder app)
        {
            var config = new HttpConfiguration();

            RegisterSwagger(config);

            RegisterWebApi(config);

            app.UseWebApi(config);
        }

        private static void RegisterSwagger(HttpConfiguration config)
        {
            config.EnableSwagger(c =>
            {
                c.SingleApiVersion("V1", "CastCrewCopyPaste Data Receiver");
            }).EnableSwaggerUi();
        }

        public static void RegisterWebApi(HttpConfiguration config)
        {
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}