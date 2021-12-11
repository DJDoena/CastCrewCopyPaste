﻿namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste.WebHost
{
    using System.Web.Http;
    using Swashbuckle.Application;

    internal sealed class SwaggerConfig
    {
        public static void RegisterOwin(HttpConfiguration config)
        {
            config.EnableSwagger(c =>
            {
                c.SingleApiVersion("V1", "CastCrewCopyPaste Data Receiver");
            }).EnableSwaggerUi();
        }
    }
}
