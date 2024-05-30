namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste
{
    using System;
    using System.Windows.Forms;
    using CastCrewCopyPaste.Resources;
    using DoenaSoft.ToolBox.Generics;
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
                    || MessageBox.Show(string.Format(MessageBoxTexts.PasteQuestion, "Cast", xmlTitle, profileTitle), MessageBoxTexts.PasteHeader, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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
                        || MessageBox.Show(string.Format(MessageBoxTexts.PasteQuestion, "Crew", xmlTitle, profileTitle), MessageBoxTexts.PasteHeader, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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

                    profile.AddCastDivider(divider.Caption.NotNull(), apiDividerType);
                }
                else if (item is CastMember cast)
                {
                    profile.AddCast(cast.FirstName.NotNull(), cast.MiddleName.NotNull(), cast.LastName.NotNull(), cast.BirthYear, cast.Role.NotNull(), cast.CreditedAs.NotNull(), cast.Voice, cast.Uncredited, cast.Puppeteer);
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

                    profile.AddCrewDivider(divider.Caption.NotNull(), apiDividerType, apiCreditType);
                }
                else if (item is CrewMember crew)
                {
                    var apiCreditType = ApiConstantsToText.GetApiCreditType(crew.CreditType);

                    var apiCreditSubtype = ApiConstantsToText.GetApiCreditSubType(crew.CreditSubtype);

                    profile.AddCrew(crew.FirstName.NotNull(), crew.MiddleName.NotNull(), crew.LastName.NotNull(), crew.BirthYear, apiCreditType, apiCreditSubtype, crew.CreditedAs.NotNull());

                    if (crew.CustomRoleSpecified)
                    {
                        profile.SetCrewCustomRoleByIndex(crewIndex, crew.CustomRole.NotNull());
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
                var information = Serializer<T>.FromString(data, CastInformation.DefaultEncoding);

                return information;
            }
            catch
            {
                return null;
            }
        }
    }
}
