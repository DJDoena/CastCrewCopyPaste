namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste
{
    using System;
    using System.Runtime.InteropServices;

    [ComVisible(false)]
    [Serializable]
    public class Settings
    {
        public DefaultValues DefaultValues;

        public string CurrentVersion;
    }

    [ComVisible(false)]
    [Serializable]
    public class DefaultValues
    {
        public bool ReceiveFromCastCrewEdit = false;
    }
}