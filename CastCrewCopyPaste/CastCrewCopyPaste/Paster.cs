using System;
using DoenaSoft.DVDProfiler.CastCrewCopyPaste.Resources;
using DoenaSoft.DVDProfiler.DVDProfilerHelper;
using DoenaSoft.DVDProfiler.DVDProfilerXML.Version400;
using Invelos.DVDProfilerPlugin;

namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste
{
    internal sealed class Paster
    {
        private IDVDProfilerAPI Api => Plugin.Api;

        public void Paste(IDVDInfo profile, string data)
        {
            var castInformation = TryGetInformationFromData<CastInformation>(data);

            if (castInformation != null)
            {
                this.PasteCast(profile, castInformation);
            }
            else
            {
                var crewInformation = TryGetInformationFromData<CrewInformation>(data);

                if (crewInformation != null)
                {
                    this.PasteCrew(profile, crewInformation);
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

            Api.SaveDVDToCollection(profile);
            Api.ReloadCurrentDVD();
            Api.UpdateProfileInListDisplay(profile.GetProfileID());
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
