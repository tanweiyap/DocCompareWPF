using ProtoBuf;

namespace DocCompareWPF.Classes
{
    [ProtoContract]
    internal class AppSettings
    {
        [ProtoMember(1)]
        public int numPanelsDragDrop = 2;

        [ProtoMember(2)]
        public string defaultFolder = "";

        [ProtoMember(3)]
        public bool isActivated = false;

        [ProtoMember(4)]
        public bool canSelectRefDoc = false;

        [ProtoMember(5)]
        public bool isProVersion = false;

        [ProtoMember(6)]
        public int maxDocCount = 5;

        [ProtoMember(7)]
        public string cultureInfo = "en-us";

        [ProtoMember(8)]
        public bool shownWalkthrough = false;

        [ProtoMember(9)]
        public bool showExtendTrial = false;

        [ProtoMember(10)]
        public bool trialExtended = false;

        [ProtoMember(11)]
        public bool skipVersion = false;

        [ProtoMember(12)]
        public string skipVersionString = "";
    }
}