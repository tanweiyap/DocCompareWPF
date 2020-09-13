using System;
using System.Collections.Generic;
using System.Text;
using ProtoBuf;

namespace DocCompareWPF.Classes
{
    [ProtoContract]
    class AppSettings
    {
        [ProtoMember (1)]
        public int numPanelsDragDrop = 3;

        [ProtoMember(2)]
        public string defaultFolder = "";

        [ProtoMember(3)]
        public bool isActivated = false;
    }
}
