using System;
using DoenaSoft.DVDProfiler.DVDProfilerXML.Version400;
using Invelos.DVDProfilerPlugin;

namespace DoenaSoft.DVDProfiler.CastCrewCopyPaste
{
    internal static class ApiConstantsToText
    {
        internal static int GetApiDividerType(DividerType dividerType)
        {
            switch (dividerType)
            {
                case DividerType.Episode:
                    {
                        return PluginConstants.DIVIDER_Episode;
                    }
                case DividerType.Group:
                    {
                        return PluginConstants.DIVIDER_Group;
                    }
                case DividerType.EndDiv:
                    {
                        return PluginConstants.DIVIDER_EndDiv;
                    }
                default:
                    {
                        throw new NotImplementedException($"Unknown divider type {dividerType}");
                    }
            }
        }

        internal static DividerType GetDividerType(int apiDividerType)
        {
            switch (apiDividerType)
            {
                case PluginConstants.DIVIDER_Episode:
                    {
                        return DividerType.Episode;
                    }
                case PluginConstants.DIVIDER_Group:
                    {
                        return DividerType.Group;
                    }
                case PluginConstants.DIVIDER_EndDiv:
                    {
                        return DividerType.EndDiv;
                    }
                default:
                    {
                        throw new NotImplementedException($"Unknown divider type {apiDividerType}");
                    }
            }
        }

        internal static int GetApiCreditType(string creditType)
        {
            switch (creditType)
            {
                case "Direction":
                    {
                        return PluginConstants.CREDIT_Direction;
                    }
                case "Writing":
                    {
                        return PluginConstants.CREDIT_Writing;
                    }
                case "Production":
                    {
                        return PluginConstants.CREDIT_Production;
                    }
                case "Cinematography":
                    {
                        return PluginConstants.CREDIT_Cinematography;
                    }
                case "Film Editing":
                    {
                        return PluginConstants.CREDIT_FilmEditing;
                    }
                case "Music":
                    {
                        return PluginConstants.CREDIT_Music;
                    }
                case "Sound":
                    {
                        return PluginConstants.CREDIT_Sound;
                    }
                case "Art":
                    {
                        return PluginConstants.CREDIT_Art;
                    }
                case "Other":
                    {
                        return 8; //there's no constant for it
                    }
                case "":
                case null:
                    {
                        return PluginConstants.CREDIT_Undefined;
                    }
                default:
                    {
                        throw new NotImplementedException($"Unknown credit type {creditType}");
                    }
            }
        }

        internal static int GetApiCreditSubType(string creditSubType)
        {
            switch (creditSubType)
            {
                case "Director":
                    {
                        return PluginConstants.CREDITSUB_Director;
                    }
                case "Original Material By":
                    {
                        return PluginConstants.CREDITSUB_OriginalMaterialBy;
                    }
                case "Screenwriter":
                    {
                        return PluginConstants.CREDITSUB_Screenwriter;
                    }
                case "Writer":
                    {
                        return PluginConstants.CREDITSUB_Writer;
                    }
                case "Original Characters By":
                    {
                        return PluginConstants.CREDITSUB_OriginalCharactersBy;
                    }
                case "Created By":
                    {
                        return PluginConstants.CREDITSUB_CreatedBy;
                    }
                case "Story By":
                    {
                        return PluginConstants.CREDITSUB_StoryBy;
                    }
                case "Producer":
                    {
                        return PluginConstants.CREDITSUB_Producer;
                    }
                case "Executive Producer":
                    {
                        return PluginConstants.CREDITSUB_ExecutiveProducer;
                    }
                case "Director of Photography":
                    {
                        return PluginConstants.CREDITSUB_DirectorOfPhotography;
                    }
                case "Cinematographer":
                    {
                        return PluginConstants.CREDITSUB_Cinematographer;
                    }
                case "Film Editor":
                    {
                        return PluginConstants.CREDITSUB_FilmEditor;
                    }
                case "Composer":
                    {
                        return PluginConstants.CREDITSUB_Composer;
                    }
                case "Song Writer":
                    {
                        return PluginConstants.CREDITSUB_SongWriter;
                    }
                case "Theme By":
                    {
                        return 2; //there's no constant for it
                    }
                case "Sound":
                    {
                        return PluginConstants.CREDITSUB_Sound;
                    }
                case "Sound Designer":
                    {
                        return PluginConstants.CREDITSUB_SoundDesigner;
                    }
                case "Supervising Sound Editor":
                    {
                        return PluginConstants.CREDITSUB_SupervisingSoundEditor;
                    }
                case "Sound Editor":
                    {
                        return PluginConstants.CREDITSUB_SoundEditor;
                    }
                case "Sound Re-Recording Mixer":
                    {
                        return PluginConstants.CREDITSUB_SoundReRecordingMixer;
                    }
                case "Production Sound Mixer":
                    {
                        return PluginConstants.CREDITSUB_ProductionSoundMixer;
                    }
                case "Production Designer":
                    {
                        return PluginConstants.CREDITSUB_ProductionDesigner;
                    }
                case "Art Director":
                    {
                        return PluginConstants.CREDITSUB_ArtDirector;
                    }
                case "Costume Designer":
                    {
                        return 2; //there's no constant for it
                    }
                case "Make-up Artist":
                    {
                        return 3; //there's no constant for it
                    }
                case "Visual Effects":
                    {
                        return 4; //there's no constant for it
                    }
                case "Make-up Effects":
                    {
                        return 5; //there's no constant for it
                    }
                case "Creature Design":
                    {
                        return 6; //there's no constant for it
                    }
                case "Custom":
                    {
                        return 254; //there's no constant for it
                    }
                default:
                    {
                        throw new NotImplementedException($"Unknown credit subtype {creditSubType}");
                    }
            }
        }

        internal static string GetCreditType(int apiCreditType)
        {
            switch (apiCreditType)
            {
                case PluginConstants.CREDIT_Direction:
                    {
                        return "Direction";
                    }
                case PluginConstants.CREDIT_Writing:
                    {
                        return "Writing";
                    }
                case PluginConstants.CREDIT_Production:
                    {
                        return "Production";
                    }
                case PluginConstants.CREDIT_Cinematography:
                    {
                        return "Cinematography";
                    }
                case PluginConstants.CREDIT_FilmEditing:
                    {
                        return "Film Editing";
                    }
                case PluginConstants.CREDIT_Music:
                    {
                        return "Music";
                    }
                case PluginConstants.CREDIT_Sound:
                    {
                        return "Sound";
                    }
                case PluginConstants.CREDIT_Art:
                    {
                        return "Art";
                    }
                case 8: //there's no constant for it
                    {
                        return "Other";
                    }
                case PluginConstants.CREDIT_Undefined:
                    {
                        return "";
                    }
                default:
                    {
                        throw new NotImplementedException($"Unknown credit type {apiCreditType}");
                    }
            }
        }

        internal static string GetCreditSubType(int apiCreditType, int apiCreditSubype)
        {
            switch (apiCreditType)
            {
                case PluginConstants.CREDIT_Direction:
                    {
                        switch (apiCreditSubype)
                        {
                            case PluginConstants.CREDITSUB_Director:
                                {
                                    return "Director";
                                }
                            case 254: //there's no constant for it
                                {
                                    return "Custom";
                                }
                            default:
                                {
                                    throw new NotImplementedException($"Unknown credit type {apiCreditType} / credit subtype {apiCreditSubype}");
                                }
                        }
                    }
                case PluginConstants.CREDIT_Writing:
                    {
                        switch (apiCreditSubype)
                        {
                            case PluginConstants.CREDITSUB_OriginalMaterialBy:
                                {
                                    return "Original Material By";
                                }
                            case PluginConstants.CREDITSUB_Screenwriter:
                                {
                                    return "Screenwriter";
                                }
                            case PluginConstants.CREDITSUB_Writer:
                                {
                                    return "Writer";
                                }
                            case PluginConstants.CREDITSUB_OriginalCharactersBy:
                                {
                                    return "Original Characters By";
                                }
                            case PluginConstants.CREDITSUB_CreatedBy:
                                {
                                    return "Created By";
                                }
                            case PluginConstants.CREDITSUB_StoryBy:
                                {
                                    return "Story By";
                                }
                            case 254: //there's no constant for it
                                {
                                    return "Custom";
                                }
                            default:
                                {
                                    throw new NotImplementedException($"Unknown credit type {apiCreditType} / credit subtype {apiCreditSubype}");
                                }
                        }
                    }
                case PluginConstants.CREDIT_Production:
                    {
                        switch (apiCreditSubype)
                        {
                            case PluginConstants.CREDITSUB_Producer:
                                {
                                    return "Producer";
                                }
                            case PluginConstants.CREDITSUB_ExecutiveProducer:
                                {
                                    return "Executive Producer";
                                }
                            case 254: //there's no constant for it
                                {
                                    return "Custom";
                                }
                            default:
                                {
                                    throw new NotImplementedException($"Unknown credit type {apiCreditType} / credit subtype {apiCreditSubype}");
                                }
                        }
                    }
                case PluginConstants.CREDIT_Cinematography:
                    {
                        switch (apiCreditSubype)
                        {
                            case PluginConstants.CREDITSUB_DirectorOfPhotography:
                                {
                                    return "Director of Photography";
                                }
                            case PluginConstants.CREDITSUB_Cinematographer:
                                {
                                    return "Cinematographer";
                                }
                            case 254: //there's no constant for it
                                {
                                    return "Custom";
                                }
                            default:
                                {
                                    throw new NotImplementedException($"Unknown credit type {apiCreditType} / credit subtype {apiCreditSubype}");
                                }
                        }
                    }
                case PluginConstants.CREDIT_FilmEditing:
                    {
                        switch (apiCreditSubype)
                        {
                            case PluginConstants.CREDITSUB_FilmEditor:
                                {
                                    return "Film Editor";
                                }
                            case 254: //there's no constant for it
                                {
                                    return "Custom";
                                }
                            default:
                                {
                                    throw new NotImplementedException($"Unknown credit type {apiCreditType} / credit subtype {apiCreditSubype}");
                                }
                        }
                    }
                case PluginConstants.CREDIT_Music:
                    {
                        switch (apiCreditSubype)
                        {
                            case PluginConstants.CREDITSUB_Composer:
                                {
                                    return "Composer";
                                }
                            case PluginConstants.CREDITSUB_SongWriter:
                                {
                                    return "Song Writer";
                                }
                            case 2: //there's no constant
                                {
                                    return "Theme By";
                                }
                            case 254: //there's no constant for it
                                {
                                    return "Custom";
                                }
                            default:
                                {
                                    throw new NotImplementedException($"Unknown credit type {apiCreditType} / credit subtype {apiCreditSubype}");
                                }
                        }
                    }
                case PluginConstants.CREDIT_Sound:
                    {
                        switch (apiCreditSubype)
                        {
                            case PluginConstants.CREDITSUB_Sound:
                                {
                                    return "Sound";
                                }
                            case PluginConstants.CREDITSUB_SoundDesigner:
                                {
                                    return "Sound Designer";
                                }
                            case PluginConstants.CREDITSUB_SupervisingSoundEditor:
                                {
                                    return "Supervising Sound Editor";
                                }
                            case PluginConstants.CREDITSUB_SoundEditor:
                                {
                                    return "Sound Editor";
                                }
                            case PluginConstants.CREDITSUB_SoundReRecordingMixer:
                                {
                                    return "Sound Re-Recording Mixer";
                                }
                            case PluginConstants.CREDITSUB_ProductionSoundMixer:
                                {
                                    return "Production Sound Mixer";
                                }
                            case 254: //there's no constant for it
                                {
                                    return "Custom";
                                }
                            default:
                                {
                                    throw new NotImplementedException($"Unknown credit type {apiCreditType} / credit subtype {apiCreditSubype}");
                                }
                        }
                    }
                case PluginConstants.CREDIT_Art:
                    {
                        switch (apiCreditSubype)
                        {
                            case PluginConstants.CREDITSUB_ProductionDesigner:
                                {
                                    return "Production Designer";
                                }
                            case PluginConstants.CREDITSUB_ArtDirector:
                                {
                                    return "Art Director";
                                }
                            case 2: //there's no constant for it
                                {
                                    return "Costume Designer";
                                }
                            case 3: //there's no constant for it
                                {
                                    return "Make-up Artist";
                                }
                            case 4: //there's no constant for it
                                {
                                    return "Visual Effects";
                                }
                            case 5: //there's no constant for it
                                {
                                    return "Make-up Effects";
                                }
                            case 6: //there's no constant for it
                                {
                                    return "Creature Design";
                                }
                            case 254: //there's no constant for it
                                {
                                    return "Custom";
                                }
                            default:
                                {
                                    throw new NotImplementedException($"Unknown credit type {apiCreditType} / credit subtype {apiCreditSubype}");
                                }
                        }
                    }
                case 8: //there's not constant for it
                    {
                        switch (apiCreditSubype)
                        {
                            case 254: //there's no constant for it
                                {
                                    return "Custom";
                                }
                            default:
                                {
                                    throw new NotImplementedException($"Unknown credit type {apiCreditType} / credit subtype {apiCreditSubype}");
                                }
                        }
                    }
                case PluginConstants.CREDIT_Undefined:
                    {
                        switch (apiCreditSubype)
                        {
                            case PluginConstants.CREDIT_Undefined:
                                {
                                    return "";
                                }
                            default:
                                {
                                    throw new NotImplementedException($"Unknown credit type {apiCreditType} / credit subtype {apiCreditSubype}");
                                }
                        }
                    }
                default:
                    {
                        throw new NotImplementedException($"Unknown credit type {apiCreditType} / credit subtype {apiCreditSubype}");
                    }
            }
        }
    }
}
