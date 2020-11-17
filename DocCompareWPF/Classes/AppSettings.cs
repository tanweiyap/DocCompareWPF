using ProtoBuf;

namespace DocCompareWPF.Classes
{
    [ProtoContract]
    class AppSettings
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
    }
}
