using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
[assembly: InternalsVisibleTo("PPTExtractor")]
namespace PowerPointTextExtractor.Models
{
    /// <summary>
    /// A record is a basic unit of data in a PowerPoint file. 
    /// It is a sequence of bytes that represents a single entity in the file.
    /// </summary>
    internal class Record
    {
        public RecordHeader Header;
        public RecordType Title => Header.Type;

        /// <summary>
        /// Creates a record that is already read.
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        internal static Record CreateRecord(BinaryReader reader)
        {
            uint version = PeekVersion(reader);
            Record record = version == 15 ? new Container() : new Atom();
            record.SetRecordHeader(reader);
            return record;
        }

        /// <summary>
        /// Will read the next 8 bytes from the reader and set the header of the record
        /// </summary>
        /// <param name="reader"></param>
        internal void SetRecordHeader(BinaryReader reader)
        {
            Header = new RecordHeader();
            var versionAndInstance = reader.ReadUInt16();//First 2 bytes
            Header.Version = versionAndInstance & 0x000FU; // First 4 bits 
            Header.Instance = (versionAndInstance & 0xFFF0U) >> 4; // Last 12 bits
            Header.TypeCode = reader.ReadUInt16();//Next 2 bytes
            Header.Size = reader.ReadUInt32();//Next 4 bytes
        }
        /// <summary>
        /// Will read the version of the Record without changing the position of the reader. 
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        internal static uint PeekVersion(BinaryReader reader)
        {
            long position = reader.BaseStream.Position;
            var versionAndInstance = reader.ReadUInt16();//First 2 bytes
            reader.BaseStream.Position = position;
            return versionAndInstance & 0x000FU;
        }
        public Record()
        {

        }
        public Record(BinaryReader reader)
        {
            Read(reader);
        }
        /// <summary>
        /// will read the Header from the record, if the header has not already been set. 
        /// </summary>
        /// <param name="reader"></param>
        public virtual void Read(BinaryReader reader)
        {
            if (Title == 0)
                SetRecordHeader(reader);
        }
    }
    public struct RecordHeader
    {
        public uint Version;
        public uint Instance;
        public ushort TypeCode;
        public RecordType Type => (RecordType)TypeCode;
        public uint Size;
        public override string ToString()
        {
            return Type.ToString();
        }
    }
    public enum RecordType
    {
        RT_TestValues = 1,
        RT_Document = 1000,
        RT_DocumentAtom = 1001,
        RT_EndDocumentAtom = 1002,
        RT_Slide = 1006,
        RT_SlideAtom = 1007,
        RT_Notes = 1008,
        RT_NotesAtom = 1009,
        RT_Environment = 1010,
        RT_SlidePersistAtom = 1011,
        RT_MainMaster = 1016,
        RT_SlideShowSlideInfoAtom = 1017,
        RT_SlideViewInfo = 1018,
        RT_GuideAtom = 1019,
        RT_ViewInfoAtom = 1021,
        RT_SlideViewInfoAtom = 1022,
        RT_VbaInfo = 1023,
        RT_VbaInfoAtom = 1024,
        RT_SlideShowDocInfoAtom = 1025,
        RT_Summary = 1026,
        RT_DocRoutingSlipAtom = 1030,
        RT_OutlineViewInfo = 1031,
        RT_SorterViewInfo = 1032,
        RT_ExternalObjectList = 1033,
        RT_ExternalObjectListAtom = 1034,
        RT_DrawingGroup = 1035,
        RT_Drawing = 1036,
        RT_GridSpacing10Atom = 1037,
        RT_RoundTripTheme12Atom = 1038,
        RT_RoundTripColorMapping12Atom = 1039,
        RT_NamedShows = 1040,
        RT_NamedShow = 1041,
        RT_NamedShowSlidesAtom = 1042,
        RT_NotesTextViewInfo9 = 1043,
        RT_NormalViewSetInfo9 = 1044,
        RT_NormalViewSetInfo9Atom = 1045,
        RT_RoundTripOriginalMainMasterId12Atom = 1052,
        RT_RoundTripCompositeMasterId12Atom = 1053,
        RT_RoundTripContentMasterInfo12Atom = 1054,
        RT_RoundTripShapeId12Atom = 1055,
        RT_RoundTripHFPlaceholder12Atom = 1056,
        RT_RoundTripContentMasterId12Atom = 1058,
        RT_RoundTripOArtTextStyles12Atom = 1059,
        RT_RoundTripHeaderFooterDefaults12Atom = 1060,
        RT_RoundTripDocFlags12Atom = 1061,
        RT_RoundTripShapeCheckSumForCL12Ato = 1062,
        RT_RoundTripNotesMasterTextStyles12Atom = 1063,
        RT_RoundTripCustomTableStyles12Atom = 1064,
        RT_List = 2000,
        RT_FontCollection = 2005,
        RT_FontCollection10 = 2006,
        RT_BookmarkCollection = 2019,
        RT_SoundCollection = 2020,
        RT_SoundCollectionAtom = 2021,
        RT_Sound = 2022,
        RT_SoundDataBlob = 2023,
        RT_BookmarkSeedAtom = 2025,
        RT_ColorSchemeAtom = 2032,
        RT_BlipCollection9 = 2040,
        RT_BlipEntity9Atom = 2041,
        RT_ExternalObjectRefAtom = 3017,
        RT_PlaceholderAtom = 3019,
        RT_ShapeAtom = 3035,
        RT_ShapeFlags10Atom = 3036,
        RT_RoundTripNewPlaceholderId12Atom = 3037,
        RT_OutlineTextRefAtom = 3998,
        RT_TextHeaderAtom = 3999,
        RT_TextCharsAtom = 4000,
        RT_StyleTextPropAtom = 4001,
        RT_MasterTextPropAtom = 4002,
        RT_TextMasterStyleAtom = 4003,
        RT_TextCharFormatExceptionAtom = 4004,
        RT_TextParagraphFormatExceptionAtom = 4005,
        RT_TextRulerAtom = 4006,
        RT_TextBookmarkAtom = 4007,
        RT_TextBytesAtom = 4008,
        RT_TextSpecialInfoDefaultAtom = 4009,
        RT_TextSpecialInfoAtom = 4010,
        RT_DefaultRulerAtom = 4011,
        RT_StyleTextProp9Atom = 4012,
        RT_TextMasterStyle9Atom = 4013,
        RT_OutlineTextProps9 = 4014,
        RT_OutlineTextPropsHeader9Atom = 4015,
        RT_TextDefaults9Atom = 4016,
        RT_StyleTextProp10Atom = 4017,
        RT_TextMasterStyle10Atom = 4018,
        RT_OutlineTextProps10 = 4019,
        RT_TextDefaults10Atom = 4020,
        RT_OutlineTextProps11 = 4021,
        RT_StyleTextProp11Atom = 4022,
        RT_FontEntityAtom = 4023,
        RT_FontEmbedDataBlob = 4024,
        RT_CString = 4026,
        RT_MetaFile = 4033,
        RT_ExternalOleObjectAtom = 4035,
        RT_Kinsoku = 4040,
        RT_Handout = 4041,
        RT_ExternalOleEmbed = 4044,
        RT_ExternalOleEmbedAtom = 4045,
        RT_ExternalOleLink = 4046,
        RT_BookmarkEntityAtom = 4048,
        RT_ExternalOleLinkAtom = 4049,
        RT_KinsokuAtom = 4050,
        RT_ExternalHyperlinkAtom = 4051,
        RT_ExternalHyperlink = 4055,
        RT_SlideNumberMetaCharAtom = 4056,
        RT_HeadersFooters = 4057,
        RT_HeadersFootersAtom = 4058,
        RT_TextInteractiveInfoAtom = 4063,
        RT_ExternalHyperlink9 = 4068,
        RT_RecolorInfoAtom = 4071,
        RT_ExternalOleControl = 4078,
        RT_SlideListWithText = 4080,
        RT_AnimationInfoAtom = 4081,
        RT_InteractiveInfo = 4082,
        RT_InteractiveInfoAtom = 4083,
        RT_UserEditAtom = 4085,
        RT_CurrentUserAtom = 4086,
        RT_DateTimeMetaCharAtom = 4087,
        RT_GenericDateMetaCharAtom = 4088,
        RT_HeaderMetaCharAtom = 4089,
        RT_FooterMetaCharAtom = 4090,
        RT_ExternalOleControlAtom = 4091,
        RT_ExternalMediaAtom = 4100,
        RT_ExternalVideo = 4101,
        RT_ExternalAviMovie = 4102,
        RT_ExternalMciMovie = 4103,
        RT_ExternalMidiAudio = 4109,
        RT_ExternalCdAudio = 4110,
        RT_ExternalWavAudioEmbedded = 4111,
        RT_ExternalWavAudioLink = 4112,
        RT_ExternalOleObjectStg = 4113,
        RT_ExternalCdAudioAtom = 4114,
        RT_ExternalWavAudioEmbeddedAtom = 4115,
        RT_AnimationInfo = 4116,
        RT_RtfDateTimeMetaCharAtom = 4117,
        RT_ExternalHyperlinkFlagsAtom = 4120,
        RT_ProgTags = 5000,
        RT_ProgStringTag = 5001,
        RT_ProgBinaryTag = 5002,
        RT_BinaryTagDataBlob = 5003,
        RT_PrintOptionsAtom = 6000,
        RT_PersistDirectoryAtom = 6002,
        RT_PresentationAdvisorFlags9Atom = 6010,
        RT_HtmlDocInfo9Atom = 6011,
        RT_HtmlPublishInfoAtom = 6012,
        RT_HtmlPublishInfo9 = 6013,
        RT_BroadcastDocInfo9 = 6014,
        RT_BroadcastDocInfo9Atom = 6015,
        RT_EnvelopeFlags9Atom = 6020,
        RT_EnvelopeData9Atom = 6021,
        RT_VisualShapeAtom = 11035,
        RT_HashCodeAtom = 11008,
        RT_VisualPageAtom = 11009,
        RT_BuildList = 10242,
        RT_BuildAtom = 10243,
        RT_ChartBuild = 11012,
        RT_ChartBuildAtom = 11013,
        RT_DiagramBuild = 11014,
        RT_DiagramBuildAtom = 10247,
        RT_ParaBuild = 10248,
        RT_ParaBuildAtom = 11017,
        RT_LevelInfoAtom = 11018,
        RT_RoundTripAnimationAtom12Atom = 11019,
        RT_RoundTripAnimationHashAtom12Atom = 11021,
        RT_Comment10 = 12000,
        RT_Comment10Atom = 12001,
        RT_CommentIndex10 = 12004,
        RT_CommentIndex10Atom = 12005,
        RT_LinkedShape10Atom = 12006,
        RT_LinkedSlide10Atom = 12007,
        RT_SlideFlags10Atom = 12010,
        RT_SlideTime10Atom = 12011,
        RT_DiffTree10 = 12012,
        RT_Diff10 = 12013,
        RT_Diff10Atom = 12014,
        RT_SlideListTableSize10Atom = 12015,
        RT_SlideListEntry10Atom = 12016,
        RT_SlideListTable10 = 12017,
        RT_CryptSession10Container = 12052,
        RT_FontEmbedFlags10Atom = 13000,
        RT_FilterPrivacyFlags10Atom = 14000,
        RT_DocToolbarStates10Atom = 14001,
        RT_PhotoAlbumInfo10Atom = 14002,
        RT_SmartTagStore11Container = 14003,
        RT_RoundTripSlideSyncInfo12 = 14068,
        OA_OfficeArtDGContainer = 0xF002,
        OA_OfficeArtFDG = 0xF008,
        OA_OfficeArtSpgrContainer = 0xF003,
        OA_OfficeArtSPContainer = 0xF004,
        OA_OfficeArtFSPGR = 0xF009,
        OA_OfficeArtFSP = 0xF00A,
        OA_OfficeArtFOPT = 0xF00B,
        RT_OfficeArtClientAnchor = 0xF010,
        RT_OfficeArtClientData = 0xF011,
        RT_OfficeArtClientTextbox = 0xF00D,
        RT_RoundTripSlideSyncInfoAtom12 = 14069,
        RT_TimeConditionContainer = 61861,
        RT_TimeNode = 61863,
        RT_TimeCondition = 61864,
        RT_TimeModifier = 61865,
        RT_TimeBehaviorContainer = 61866,
        RT_TimeAnimateBehaviorContainer = 61867,
        RT_TimeColorBehaviorContainer = 61868,
        RT_TimeEffectBehaviorContainer = 61869,
        RT_TimeMotionBehaviorContainer = 61870,
        RT_TimeRotationBehaviorContainer = 61871,
        RT_TimeScaleBehaviorContainer = 61872,
        RT_TimeSetBehaviorContainer = 61873,
        RT_TimeCommandBehaviorContainer = 61874,
        RT_TimeBehavior = 61875,
        RT_TimeAnimateBehavior = 61876,
        RT_TimeColorBehavior = 61877,
        RT_TimeEffectBehavior = 61878,
        RT_TimeMotionBehavior = 61879,
        RT_TimeRotationBehavior = 61880,
        RT_TimeScaleBehavior = 61881,
        RT_TimeSetBehavior = 61882,
        RT_TimeCommandBehavior = 61883,
        RT_TimeClientVisualElement = 61884,
        RT_TimePropertyList = 61885,
        RT_TimeVariantList = 61886,
        RT_TimeAnimationValueList = 61887,
        RT_TimeIterateData = 61888,
        RT_TimeSequenceData = 61889,
        RT_TimeVariant = 61890,
        RT_TimeAnimationValue = 61891,
        RT_TimeExtTimeNodeContainer = 61892,
        RT_TimeSubEffectContainer = 61893,



    };
}
