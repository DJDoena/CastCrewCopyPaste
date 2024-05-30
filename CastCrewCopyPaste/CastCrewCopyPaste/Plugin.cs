namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using CastCrewCopyPaste.Resources;
    using DoenaSoft.ToolBox.Generics;
    using DVDProfilerHelper;
    using DVDProfilerXML.Version400;
    using Invelos.DVDProfilerPlugin;

    [ComVisible(true)]
    [Guid(ClassGuid.ClassID)]
    public class Plugin : IDVDProfilerPlugin, IDVDProfilerPluginInfo
    {
        private readonly string _applicationPath;

        private readonly string _errorFile;

        private readonly string _settingsFile;

        private const int CopyCastMenuId = 1;

        private string _copyCastMenuToken = "";

        private const int CopyCrewMenuId = 2;

        private string _copyCrewMenuToken = "";

        private const int PasteMenuId = 3;

        private string _pasteMenuToken = "";

        private const int ReceiverSettingMenuId = 4;

        private string _receiverSettingMenuToken = "";

        private CastCrewReceiverService _wcfService;

        private Settings _settings;

        internal static IDVDProfilerAPI Api { get; set; }

        private AssemblyName Assembly => System.Reflection.Assembly.GetAssembly(this.GetType()).GetName();

        private Version PluginVersion => this.Assembly.Version;

        public Plugin()
        {
            _applicationPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Doena Soft\CastCrewCopyPasteSettings\";

            _errorFile = Environment.GetEnvironmentVariable("TEMP") + @"\CastCrewCopyPasteCrash.xml";

            _settingsFile = _applicationPath + "CastCrewCopyPasteSettings.xml";
        }

        #region I... Members

        #region IDVDProfilerPlugin

        public void Load(IDVDProfilerAPI api)
        {
            //Debugger.Launch();

            try
            {
                Api = api;

                if (Directory.Exists(_applicationPath) == false)
                {
                    Directory.CreateDirectory(_applicationPath);
                }

                if (File.Exists(_settingsFile))
                {
                    try
                    {
                        _settings = Serializer<Settings>.Deserialize(_settingsFile);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(string.Format(MessageBoxTexts.FileCantBeRead, _settingsFile, ex.Message), MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                this.CreateSettings();

                if (_settings.DefaultValues.ReceiveFromCastCrewEdit)
                {
                    this.LoadReceiver();
                }

                if (Directory.Exists(_applicationPath) == false)
                {
                    Directory.CreateDirectory(_applicationPath);
                }

                Api.RegisterForEvent(PluginConstants.EVENTID_FormCreated);

                _copyCastMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form, @"DVD", "Copy Cast", CopyCastMenuId);

                _copyCrewMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form, @"DVD", "Copy Crew", CopyCrewMenuId);

                _pasteMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form, @"DVD", "Paste Cast / Crew", PasteMenuId);

                _receiverSettingMenuToken = Api.RegisterMenuItem(PluginConstants.FORMID_Main, PluginConstants.MENUID_Form, @"Tools", "Enable Cast/Crew Edit 2 Receiver", ReceiverSettingMenuId);

                api.SetRegisteredMenuItemChecked(_receiverSettingMenuToken, _settings.DefaultValues.ReceiveFromCastCrewEdit);

                var pluginVersion = this.PluginVersion.ToString();

                if (_settings.CurrentVersion != pluginVersion)
                {
                    this.OpenReadme();

                    _settings.CurrentVersion = pluginVersion;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadReceiver()
        {
            _wcfService = new CastCrewReceiverService();

            _wcfService.Start();
        }

        public void Unload()
        {
            this.UnloadReceiver();

            try
            {
                Serializer<Settings>.Serialize(_settingsFile, _settings);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(MessageBoxTexts.FileCantBeWritten, _settingsFile, ex.Message), MessageBoxTexts.ErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Api.UnregisterMenuItem(_copyCastMenuToken);
            Api.UnregisterMenuItem(_copyCrewMenuToken);
            Api.UnregisterMenuItem(_pasteMenuToken);
            Api.UnregisterMenuItem(_receiverSettingMenuToken);

            Api = null;
        }

        private void UnloadReceiver()
        {
            _wcfService?.Finish();
            _wcfService = null;
        }

        public void HandleEvent(int EventType, object EventData)
        {
            try
            {
                if (EventType == PluginConstants.EVENTID_CustomMenuClick)
                {
                    this.HandleMenuClick((int)EventData);
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

                    this.LogException(ex);
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

        public int GetVersionMajor() => this.PluginVersion.Major;

        public int GetVersionMinor() => this.PluginVersion.Minor * 100 + this.PluginVersion.Build * 10 + this.PluginVersion.Revision;

        #endregion

        #endregion

        private void HandleMenuClick(int MenuEventID)
        {
            switch (MenuEventID)
            {
                case CopyCastMenuId:
                    {
                        this.TryCopyPaste(this.CopyCast);

                        break;
                    }
                case CopyCrewMenuId:
                    {
                        this.TryCopyPaste(this.CopyCrew);

                        break;
                    }
                case PasteMenuId:
                    {
                        this.TryCopyPaste(this.Paste);

                        break;
                    }
                case ReceiverSettingMenuId:
                    {
                        _settings.DefaultValues.ReceiveFromCastCrewEdit = !_settings.DefaultValues.ReceiveFromCastCrewEdit;

                        Api.SetRegisteredMenuItemChecked(_receiverSettingMenuToken, _settings.DefaultValues.ReceiveFromCastCrewEdit);

                        if (_settings.DefaultValues.ReceiveFromCastCrewEdit)
                        {
                            this.LoadReceiver();
                        }
                        else
                        {
                            this.UnloadReceiver();
                        }

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
                        FirstName = firstName.NotNull(),
                        MiddleName = middleName.NotNull(),
                        LastName = lastName.NotNull(),
                        BirthYear = birthYear,
                        Role = role.NotNull(),
                        CreditedAs = creditedAs.NotNull(),
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
                        Caption = caption.NotNull(),
                        Type = dividerType,
                    });
                }
            }

            var castInformation = new CastInformation()
            {
                Title = profile.GetTitle().NotNull(),
                CastList = castList.ToArray(),
            };

            var xml = Serializer<CastInformation>.ToString(castInformation, CastInformation.DefaultEncoding);

            try
            {
                Clipboard.SetDataObject(xml, true, 4, 250);
            }
            catch (Exception ex) //clipboard in use
            {
                MessageBox.Show(ex.Message, MessageBoxTexts.CriticalErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                        FirstName = firstName.NotNull(),
                        MiddleName = middleName.NotNull(),
                        LastName = lastName.NotNull(),
                        BirthYear = birthYear,
                        CreditType = creditType,
                        CreditSubtype = creditSubtype,
                        CustomRole = customRole,
                        CustomRoleSpecified = !string.IsNullOrWhiteSpace(customRole),
                        CreditedAs = creditedAs.NotNull(),
                    });
                }
                else
                {
                    profile.GetCrewDividerByIndex(crewIndex, out var caption, out var apiDividerType, out apiCreditType);

                    var dividerType = ApiConstantsToText.GetDividerType(apiDividerType);

                    var creditType = ApiConstantsToText.GetCreditType(apiCreditType);

                    crewList.Add(new CrewDivider()
                    {
                        Caption = caption.NotNull(),
                        Type = dividerType,
                        CreditType = creditType,
                    });
                }
            }

            var crewInformation = new CrewInformation()
            {
                Title = profile.GetTitle().NotNull(),
                CrewList = crewList.ToArray(),
            };

            var xml = Serializer<CrewInformation>.ToString(crewInformation, CrewInformation.DefaultEncoding);

            try
            {
                Clipboard.SetDataObject(xml, true, 4, 250);
            }
            catch (Exception ex) //clipboard in use
            {
                MessageBox.Show(ex.Message, MessageBoxTexts.CriticalErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Paste(IDVDInfo profile)
        {
            try
            {
                (new Paster()).Paste(profile, Clipboard.GetText());
            }
            catch (PasteException ex)
            {
                MessageBox.Show(ex.Message, MessageBoxTexts.WarningHeader, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
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
            catch (COMException)
            {
                throw;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, MessageBoxTexts.CriticalErrorHeader, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LogException(Exception ex)
        {
            ex = this.WrapCOMException(ex);

            var exceptionXml = new ExceptionXml(ex);

            Serializer<ExceptionXml>.Serialize(_errorFile, exceptionXml);
        }

        private Exception WrapCOMException(Exception ex)
        {
            var returnEx = ex;

            if (ex is COMException comEx)
            {
                var lastApiError = Api.GetLastError();

                returnEx = new EnhancedCOMException(lastApiError, comEx);
            }

            return returnEx;
        }

        private void CreateSettings()
        {
            if (_settings == null)
            {
                _settings = new Settings();
            }

            if (_settings.DefaultValues == null)
            {
                _settings.DefaultValues = new DefaultValues();
            }
        }

        private void OpenReadme()
        {
            //System.Diagnostics.Debugger.Launch();

            var dllPath = new FileInfo((new Uri(this.Assembly.CodeBase)).LocalPath);

            var helpFile = Path.Combine(dllPath.DirectoryName, "ReadMe", "ReadMe.html");

            if (File.Exists(helpFile))
            {
                using (var helpForm = new HelpForm(helpFile))
                {
                    helpForm.Text = "Read Me";
                    helpForm.ShowDialog();
                }
            }
        }

        #region Plugin Registering

        [DllImport("user32.dll")]
        public extern static int SetParent(int child, int parent);

        [ComImport, Guid("0002E005-0000-0000-C000-000000000046")]
        internal class StdComponentCategoriesMgr { }

        [ComRegisterFunction]
        public static void RegisterServer(Type _)
        {
            var cr = (CategoryRegistrar.ICatRegister)new StdComponentCategoriesMgr();

            var clsidThis = new Guid(ClassGuid.ClassID);

            var catid = new Guid("833F4274-5632-41DB-8FC5-BF3041CEA3F1");

            cr.RegisterClassImplCategories(ref clsidThis, 1, new[] { catid });
        }

        [ComUnregisterFunction]
        public static void UnregisterServer(Type _)
        {
            var cr = (CategoryRegistrar.ICatRegister)new StdComponentCategoriesMgr();

            var clsidThis = new Guid(ClassGuid.ClassID);

            var catid = new Guid("833F4274-5632-41DB-8FC5-BF3041CEA3F1");

            cr.UnRegisterClassImplCategories(ref clsidThis, 1, new[] { catid });
        }

        #endregion
    }
}