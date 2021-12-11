namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste
{
    using System;
    using System.Windows.Forms;
    using CastCrewCopyPaste.Resources;
    using DVDProfilerHelper;
    using DVDProfilerXML.Version400;
    using Invelos.DVDProfilerPlugin;

    internal sealed class Paster
    {
        private IDVDProfilerAPI Api => Plugin.Api;

        public void Paste(IDVDInfo profile, string xml)
        {
            var profileTitle = profile.GetTitle();

            var profileTitleWithYear = $"{profileTitle} ({profile.GetProductionYear()})";

            var castInformation = TryGetInformationFromData<CastInformation>(xml);

            if (castInformation != null)
            {
                var xmlTitle = castInformation.Title;

                if (string.Equals(profileTitle, xmlTitle, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(profileTitleWithYear, xmlTitle, StringComparison.OrdinalIgnoreCase)
                    || MessageBox.Show(string.Format(MessageBoxTexts.PasteQuestion, xmlTitle, profileTitle), MessageBoxTexts.PasteHeader, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    this.PasteCast(profile, castInformation);
                }
            }
            else
            {
                var crewInformation = TryGetInformationFromData<CrewInformation>(xml);

                if (crewInformation != null)
                {
                    var xmlTitle = crewInformation.Title;

                    if (string.Equals(profileTitle, xmlTitle, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(profileTitleWithYear, xmlTitle, StringComparison.OrdinalIgnoreCase)
                        || MessageBox.Show(string.Format(MessageBoxTexts.PasteQuestion, xmlTitle, profileTitle), MessageBoxTexts.PasteHeader, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        this.PasteCrew(profile, crewInformation);
                    }
                }
                else
                {
                    throw new PasteException(MessageBoxTexts.UnknownInformationInClipboard);
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

                    profile.AddCastDivider(Plugin.NotNull(divider.Caption), apiDividerType);
                }
                else if (item is CastMember cast)
                {
                    profile.AddCast(Plugin.NotNull(cast.FirstName), Plugin.NotNull(cast.MiddleName), Plugin.NotNull(cast.LastName), cast.BirthYear, Plugin.NotNull(cast.Role), Plugin.NotNull(cast.CreditedAs), cast.Voice, cast.Uncredited, cast.Puppeteer);
                }
                else
                {
                    throw new NotImplementedException($"Unknown crew item {item}");
                }
            }

            this.Api.SaveDVDToCollection(profile);
            this.Api.ReloadCurrentDVD();
            this.Api.UpdateProfileInListDisplay(profile.GetProfileID());
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

                    profile.AddCrewDivider(Plugin.NotNull(divider.Caption), apiDividerType, apiCreditType);
                }
                else if (item is CrewMember crew)
                {
                    var apiCreditType = ApiConstantsToText.GetApiCreditType(crew.CreditType);

                    var apiCreditSubtype = ApiConstantsToText.GetApiCreditSubType(crew.CreditSubtype);

                    profile.AddCrew(Plugin.NotNull(crew.FirstName), Plugin.NotNull(crew.MiddleName), Plugin.NotNull(crew.LastName), crew.BirthYear, apiCreditType, apiCreditSubtype, Plugin.NotNull(crew.CreditedAs));

                    if (crew.CustomRoleSpecified)
                    {
                        profile.SetCrewCustomRoleByIndex(crewIndex, Plugin.NotNull(crew.CustomRole));
                    }
                }
                else
                {
                    throw new NotImplementedException($"Unknown crew item {item}");
                }
            }

            this.Api.SaveDVDToCollection(profile);
            this.Api.ReloadCurrentDVD();
            this.Api.UpdateProfileInListDisplay(profile.GetProfileID());
        }

        private static T TryGetInformationFromData<T>(string data) where T : class, new()
        {
            try
            {
                var information = DVDProfilerSerializer<T>.FromString(data, CastInformation.DefaultEncoding);

                return information;
            }
            catch
            {
                return null;
            }
        }
    }
}
