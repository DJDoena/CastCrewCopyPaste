using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DoenaSoft.DVDProfiler.CastCrewCopyPaste.Resources;
using DoenaSoft.DVDProfiler.DVDProfilerHelper;
using DoenaSoft.DVDProfiler.DVDProfilerXML.Version400;
using Invelos.DVDProfilerPlugin;

namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste
{
    [ComVisible(true)]
    [Guid(ClassGuid.ClassID)]
    public class Plugin : IDVDProfilerPlugin, IDVDProfilerPluginInfo
    {
        private readonly string _errorFile;

        private readonly string _applicationPath;

        private readonly Version _pluginVersion;

        private IDVDProfilerAPI Api { get; set; }

        private const int CopyCastMenuId = 1;

        private string _copyCastMenuToken = "";

        private const int CopyCrewMenuId = 2;

        private string _copyCrewMenuToken = "";

        private const int PasteMenuId = 3;

        private string _pasteMenuToken = "";


        public Plugin()
        {
            _applicationPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Doena Soft\CastCrewCopyPasteSettings\";

            _errorFile = Environment.GetEnvironmentVariable("TEMP") + @"\CastCrewCopyPasteCrash.xml";

            _pluginVersion = System.Reflection.Assembly.GetAssembly(GetType()).GetName().Version;
        }

        #region I... Members

        #region IDVDProfilerPlugin

        public void Load(IDVDProfilerAPI api)
        {
            Api = api;

            if (Directory.Exists(_applicationPath) == false)
            {
                Directory.CreateDirectory(_applicationPath);
            }

            Api.RegisterForEvent(PluginConstants.EVENTID_FormCreated);

            _copyCastMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                , @"DVD", "Copy Cast", CopyCastMenuId);
            _copyCrewMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                , @"DVD", "Copy Crew", CopyCrewMenuId);
            _pasteMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form
                , @"DVD", "Paste Cast / Crew", PasteMenuId);
        }

        public void Unload()
        {
            Api.UnregisterMenuItem(_copyCastMenuToken);
            Api.UnregisterMenuItem(_copyCrewMenuToken);
            Api.UnregisterMenuItem(_pasteMenuToken);

            Api = null;
        }

        public void HandleEvent(int EventType, object EventData)
        {
            try
            {
                if (EventType == PluginConstants.EVENTID_CustomMenuClick)
                {
                    HandleMenuClick((int)EventData);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    MessageBox.Show(string.Format(MessageBoxTexts.CriticalError, ex.Message, _errorFile), MessageBoxTexts.CriticalErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Stop);

                    if (File.Exists(_errorFile))
                    {
                        File.Delete(_errorFile);
                    }

                    LogException(ex);
                }
                catch (Exception inEx)
                {
                    MessageBox.Show(string.Format(MessageBoxTexts.FileCantBeWritten, _errorFile, inEx.Message), MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        #endregion

        #region IDVDProfilerPluginInfo

        public string GetName() => Texts.PluginName;

        public string GetDescription() => Texts.PluginDescription;

        public string GetAuthorName() => "Doena Soft.";

        public string GetAuthorWebsite() => Texts.PluginUrl;

        public int GetPluginAPIVersion() => PluginConstants.API_VERSION;

        public int GetVersionMajor() => _pluginVersion.Major;

        public int GetVersionMinor() => _pluginVersion.Minor;

        #endregion

        #endregion

        private void HandleMenuClick(int MenuEventID)
        {
            switch (MenuEventID)
            {
                case CopyCastMenuId:
                    {
                        TryCopyPaste(CopyCast);

                        break;
                    }
                case CopyCrewMenuId:
                    {
                        TryCopyPaste(CopyCrew);

                        break;
                    }
                case PasteMenuId:
                    {
                        TryCopyPaste(Paste);

                        break;
                    }
            }
        }

        private void CopyCast(IDVDInfo profile)
        {
            var castCount = profile.GetCastCount();

            var castList = new List<object>(castCount);

            for (var castIndex = 0; castIndex < castCount; castIndex++)
            {
                profile.GetCastByIndex(castIndex, out var firstName, out var middleName, out var lastName, out var birthYear, out var role, out var creditedAs, out var voice, out var uncredited, out var puppeteer);

                if (firstName != null)
                {
                    castList.Add(new CastMember()
                    {
                        FirstName = NotNull(firstName),
                        MiddleName = NotNull(middleName),
                        LastName = NotNull(lastName),
                        BirthYear = birthYear,
                        Role = NotNull(role),
                        CreditedAs = NotNull(creditedAs),
                        Voice = voice,
                        Uncredited = uncredited,
                        Puppeteer = puppeteer,
                    });
                }
                else
                {
                    profile.GetCastDividerByIndex(castIndex, out var caption, out var apiDividerType);

                    var dividerType = ApiConstantsToText.GetDividerType(apiDividerType);

                    castList.Add(new Divider()
                    {
                        Caption = NotNull(caption),
                        Type = dividerType,
                    });
                }
            }

            var castInformation = new CastInformation()
            {
                Title = NotNull(profile.GetTitle()),
                CastList = castList.ToArray(),
            };

            var text = DVDProfilerSerializer<CastInformation>.ToString(castInformation, CastInformation.DefaultEncoding);

            Clipboard.SetText(text);
        }

        private void CopyCrew(IDVDInfo profile)
        {
            var crewCount = profile.GetCrewCount();

            var crewList = new List<object>(crewCount);

            for (var crewIndex = 0; crewIndex < crewCount; crewIndex++)
            {
                profile.GetCrewByIndex(crewIndex, out var firstName, out var middleName, out var lastName, out var birthYear, out var apiCreditType, out var apiCreditSubtype, out var creditedAs);

                if (firstName != null)
                {
                    var customRole = profile.GetCrewCustomRoleByIndex(crewIndex);

                    var creditType = ApiConstantsToText.GetCreditType(apiCreditType);

                    var creditSubtype = ApiConstantsToText.GetCreditSubType(apiCreditType, apiCreditSubtype);

                    crewList.Add(new CrewMember()
                    {
                        FirstName = NotNull(firstName),
                        MiddleName = NotNull(middleName),
                        LastName = NotNull(lastName),
                        BirthYear = birthYear,
                        CreditType = creditType,
                        CreditSubtype = creditSubtype,
                        CustomRole = customRole,
                        CustomRoleSpecified = !string.IsNullOrEmpty(customRole),
                        CreditedAs = NotNull(creditedAs),
                    });
                }
                else
                {
                    profile.GetCrewDividerByIndex(crewIndex, out var caption, out var apiDividerType, out apiCreditType);

                    var dividerType = ApiConstantsToText.GetDividerType(apiDividerType);

                    var creditType = ApiConstantsToText.GetCreditType(apiCreditType);

                    crewList.Add(new CrewDivider()
                    {
                        Caption = NotNull(caption),
                        Type = dividerType,
                        CreditType = creditType,
                    });
                }
            }

            var crewInformation = new CrewInformation()
            {
                Title = NotNull(profile.GetTitle()),
                CrewList = crewList.ToArray(),
            };

            var text = DVDProfilerSerializer<CrewInformation>.ToString(crewInformation, CrewInformation.DefaultEncoding);

            Clipboard.SetText(text);
        }

        private void Paste(IDVDInfo profile)
        {
            var castInformation = TryGetInformationFromClipboard<CastInformation>();

            if (castInformation != null)
            {
                PasteCast(profile, castInformation);
            }
            else
            {
                var crewInformation = TryGetInformationFromClipboard<CrewInformation>();

                if (crewInformation != null)
                {
                    PasteCrew(profile, crewInformation);
                }
                else
                {
                    MessageBox.Show(MessageBoxTexts.UnknownInformationInClipboard, MessageBoxTexts.WarningHeader, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void PasteCast(IDVDInfo profile, CastInformation castInformation)
        {
            profile.ClearCast();

            for (var castIndex = 0; castIndex < (castInformation.CastList?.Length ?? 0); castIndex++)
            {
                var item = castInformation.CastList[castIndex];

                if (item is Divider divider)
                {
                    var apiDividerType = ApiConstantsToText.GetApiDividerType(divider.Type);

                    profile.AddCastDivider(NotNull(divider.Caption), apiDividerType);
                }
                else if (item is CastMember cast)
                {
                    profile.AddCast(NotNull(cast.FirstName), NotNull(cast.MiddleName), NotNull(cast.LastName), cast.BirthYear, NotNull(cast.Role), NotNull(cast.CreditedAs), cast.Voice, cast.Uncredited, cast.Puppeteer);
                }
                else
                {
                    throw new NotImplementedException($"Unknown crew item {item}");
                }
            }

            Api.SaveDVDToCollection(profile);
            Api.ReloadCurrentDVD();
            Api.UpdateProfileInListDisplay(profile.GetProfileID());
        }

        private void PasteCrew(IDVDInfo profile, CrewInformation crewInformation)
        {
            profile.ClearCrew();

            for (var crewIndex = 0; crewIndex < (crewInformation.CrewList?.Length ?? 0); crewIndex++)
            {
                var item = crewInformation.CrewList[crewIndex];

                if (item is CrewDivider divider)
                {
                    var apiDividerType = ApiConstantsToText.GetApiDividerType(divider.Type);

                    var apiCreditType = ApiConstantsToText.GetApiCreditType(divider.CreditType);

                    profile.AddCrewDivider(NotNull(divider.Caption), apiDividerType, apiCreditType);
                }
                else if (item is CrewMember crew)
                {
                    var apiCreditType = ApiConstantsToText.GetApiCreditType(crew.CreditType);

                    var apiCreditSubtype = ApiConstantsToText.GetApiCreditSubType(crew.CreditSubtype);

                    profile.AddCrew(NotNull(crew.FirstName), NotNull(crew.MiddleName), NotNull(crew.LastName), crew.BirthYear, apiCreditType, apiCreditSubtype, NotNull(crew.CreditedAs));

                    if (crew.CustomRoleSpecified)
                    {
                        profile.SetCrewCustomRoleByIndex(crewIndex, NotNull(crew.CustomRole));
                    }
                }
                else
                {
                    throw new NotImplementedException($"Unknown crew item {item}");
                }
            }

            Api.SaveDVDToCollection(profile);
            Api.ReloadCurrentDVD();
            Api.UpdateProfileInListDisplay(profile.GetProfileID());
        }

        private void TryCopyPaste(Action<IDVDInfo> action)
        {
            try
            {
                var currentDisplayedProfileId = Api.GetDisplayedDVD()?.GetProfileID();

                if (!string.IsNullOrEmpty(currentDisplayedProfileId))
                {
                    Api.DVDByProfileID(out var profile, currentDisplayedProfileId, -1, -1);

                    action(profile);
                }
                else
                {
                    MessageBox.Show(MessageBoxTexts.NoProfileSelected, MessageBoxTexts.WarningHeader, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ex = WrapCOMException(ex);

                MessageBox.Show(ex.Message, MessageBoxTexts.CriticalErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static T TryGetInformationFromClipboard<T>() where T : class, new()
        {
            try
            {
                var information = DVDProfilerSerializer<T>.FromString(Clipboard.GetText(), CastInformation.DefaultEncoding);

                return information;
            }
            catch
            {
                return null;
            }
        }

        private static string NotNull(string text) => text ?? string.Empty;

        private void LogException(Exception ex)
        {
            ex = WrapCOMException(ex);

            var exceptionXml = new ExceptionXml(ex);

            DVDProfilerSerializer<ExceptionXml>.Serialize(_errorFile, exceptionXml);
        }

        private Exception WrapCOMException(Exception ex)
        {
            var returnEx = ex;

            if (ex is COMException comEx)
            {
                var lastApiError = Api.GetLastError();

                returnEx = new EnhancedCOMException(comEx, lastApiError);
            }

            return returnEx;
        }

        #region Plugin Registering

        [DllImport("user32.dll")]
        public extern static int SetParent(int child, int parent);

        [ComImport(), Guid("0002E005-0000-0000-C000-000000000046")]
        internal class StdComponentCategoriesMgr { }

        [ComRegisterFunction()]
        public static void RegisterServer(Type _)
        {
            CategoryRegistrar.ICatRegister cr = (CategoryRegistrar.ICatRegister)new StdComponentCategoriesMgr();

            Guid clsidThis = new Guid(ClassGuid.ClassID);

            Guid catid = new Guid("833F4274-5632-41DB-8FC5-BF3041CEA3F1");

            cr.RegisterClassImplCategories(ref clsidThis, 1, new Guid[] { catid });
        }

        [ComUnregisterFunction()]
        public static void UnregisterServer(Type _)
        {
            CategoryRegistrar.ICatRegister cr = (CategoryRegistrar.ICatRegister)new StdComponentCategoriesMgr();

            Guid clsidThis = new Guid(ClassGuid.ClassID);

            Guid catid = new Guid("833F4274-5632-41DB-8FC5-BF3041CEA3F1");

            cr.UnRegisterClassImplCategories(ref clsidThis, 1, new Guid[] { catid });
        }

        #endregion
    }
}