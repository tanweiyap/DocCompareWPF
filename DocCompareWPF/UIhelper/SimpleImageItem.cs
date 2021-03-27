using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace DocCompareWPF.UIhelper
{
    public class SimpleImageItem : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public Thickness Margin { get; set; }

        private string _pathToFile;
        private string _pathToFileHidden;
        private Visibility _eodVisi;
        private double _blurRadius;
        private Visibility _showHidden;
        private Visibility _showPPTNoteButton;

        public string ShowPPTSpeakerNotesButtonName { get; set; }
        public string HiddenPPTGridName { get; set; }
        public string PPTSpeakerNoteGridName { get; set; }
        public string ClosePPTSpeakerNotesButtonName { get; set; }
        public string PPTSpeakerNotes { get; set; }

        public string ShowSpeakerNotesTooltip { get; set; }
        public bool ShowSpeakerNoteEnable { get; set; }


        public string PathToFile
        {
            get
            {
                return _pathToFile;
            }

            set
            {
                _pathToFile = value;
                OnPropertyChanged();
            }
        }

        public string PathToFileHidden
        {
            get
            {
                return _pathToFileHidden;
            }

            set
            {
                _pathToFileHidden = value;
                OnPropertyChanged();
            }
        }

        public Visibility EoDVisi
        {
            get
            {
                return _eodVisi;
            }

            set
            {
                _eodVisi = value;
                OnPropertyChanged();
            }
        }

        public Visibility showHidden
        {
            get
            {
                return _showHidden;
            }

            set
            {
                _showHidden = value;
                OnPropertyChanged();
            }
        }
        public Visibility showPPTSpeakerNotesButton
        {
            get
            {
                return _showPPTNoteButton;
            }

            set
            {
                _showPPTNoteButton = value;
                OnPropertyChanged();
            }
        }

        public double BlurRadius
        {
            get
            {
                return _blurRadius;
            }

            set
            {
                _blurRadius = value;
                OnPropertyChanged();
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
