namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste
{
    using System;
    using System.Runtime.InteropServices;
    using System.ServiceModel;
    using Invelos.DVDProfilerPlugin;
    using Resources;

    public sealed class CastCrewReceiverService : ICastCrewReceiver
    {
        private IDVDProfilerAPI Api => Plugin.Api;

        private ServiceHost _serviceHost;

        public void Start()
        {
            _serviceHost = new ServiceHost(typeof(CastCrewReceiverService));

            var binding = new NetTcpBinding();

            binding.Security.Mode = SecurityMode.None;
            binding.MaxBufferSize = int.MaxValue;
            binding.MaxReceivedMessageSize = int.MaxValue;

            _serviceHost.AddServiceEndpoint(typeof(ICastCrewReceiver), binding, new Uri(CastCrewReceiverServiceContract.TcpAddress));

            _serviceHost.Open();
        }

        public void Finish()
        {
            _serviceHost.Close();
            ((IDisposable)_serviceHost).Dispose();
        }

        public string Receive(string xml)
        {
            try
            {
                var currentDisplayedProfileId = this.Api.GetDisplayedDVD()?.GetProfileID();

                if (!string.IsNullOrEmpty(currentDisplayedProfileId))
                {
                    this.Api.DVDByProfileID(out var profile, currentDisplayedProfileId, -1, -1);

                    if (!string.IsNullOrWhiteSpace(xml))
                    {
                        (new Paster()).Paste(profile, xml);
                    }

                    return "OK";
                }
                else
                {
                    return MessageBoxTexts.NoProfileSelected;
                }
            }
            catch (COMException ex)
            {
                return $"COM Exception, ErrorCode: {ex.ErrorCode}, Message: {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"Exception, Message: {ex.Message}";
            }
        }
    }
}