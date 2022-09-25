//namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste.WebHost.Controllers
//{
//    using System;
//    using System.Runtime.InteropServices;
//    using System.Web.Http;
//    using System.Web.Http.Description;
//    using CastCrewCopyPaste.Resources;
//    using Invelos.DVDProfilerPlugin;

//    [RoutePrefix("api/Receiver")]
//    public sealed class ReceiverController : ApiController
//    {
//        private IDVDProfilerAPI Api => Plugin.Api;

//        [HttpPost]
//        [Route(nameof(Receive))]
//        [ResponseType(typeof(void))]
//        public IHttpActionResult Receive([FromBody] string xml)
//        {
//            try
//            {
//                var currentDisplayedProfileId = this.Api.GetDisplayedDVD()?.GetProfileID();

//                if (!string.IsNullOrEmpty(currentDisplayedProfileId))
//                {
//                    this.Api.DVDByProfileID(out var profile, currentDisplayedProfileId, -1, -1);

//                    if (!string.IsNullOrWhiteSpace(xml))
//                    {
//                        (new Paster()).Paste(profile, xml);
//                    }

//                    return this.Ok();
//                }
//                else
//                {
//                    return this.BadRequest(MessageBoxTexts.NoProfileSelected);
//                }
//            }
//            catch (COMException ex)
//            {
//                return this.BadRequest($"COM Exception, ErrorCode: {ex.ErrorCode}, Message: {ex.Message}");
//            }
//            catch (Exception ex)
//            {
//                return this.BadRequest($"Exception, Message: {ex.Message}");
//            }
//        }
//    }
//}