using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Media;

namespace DocCompareWPF.UIhelper
{
    public class CompareMainItem : INotifyPropertyChanged
    {
        private Visibility _showMask;

        public string Document1 { get; set; }
        public string Document2 { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public bool AniDiffButtonEnable { get; set; }
        public string AnimateDiffLeftButtonName { get; set; }
        public string AnimateDiffRightButtonName { get; set; }
        public string ImgAniLeftName { get; set; }
        public string ImgAniRightName { get; set; }
        public string ImgGridLeftName { get; set; }
        public string ImgGridName { get; set; }
        public string ImgGridRightName { get; set; }
        public string ImgLeftName { get; set; }
        public string ImgMaskRightName { get; set; }
        public string ImgRightName { get; set; }
        public Thickness Margin { get; set; }

        private string _pathToAniImgLeft;
        private string _pathToAniImgRight;
        private string _pathToImgLeft;
        private string _pathToImgRight;
        private string _pathToMaskImgRight;
        public string PathToAniImgLeft { get { return _pathToAniImgLeft; } set { _pathToAniImgLeft = value; OnPropertyChanged(); } }
        public string PathToAniImgRight { get { return _pathToAniImgRight; } set { _pathToAniImgRight = value; OnPropertyChanged(); } }
        public string PathToImgLeft { get { return _pathToImgLeft; } set { _pathToImgLeft = value; OnPropertyChanged(); } }
        public string PathToImgRight { get { return _pathToImgRight; } set { _pathToImgRight = value; OnPropertyChanged(); } }
        public string PathToMaskImgRight { get { return _pathToMaskImgRight; } set { _pathToMaskImgRight = value; OnPropertyChanged(); } }

        private double _blurRadiusLeft;
        private double _blurRadiusRight;
        public string ShowSpeakerNotesTooltip { get; set; }
        public bool ShowSpeakerNoteEnable { get; set; }

        private Visibility _showHiddenLeft;
        private Visibility _showHiddenRight;
        private Visibility _pptNoteGridLeftVisi;
        private Visibility _pptNoteGridRightVisi;

        public string PPTSpeakerNoteGridNameLeft { get; set; }
        public string PPTSpeakerNoteGridNameRight { get; set; }
        public string ClosePPTSpeakerNotesButtonNameLeft { get; set; }
        public string ClosePPTSpeakerNotesButtonNameRight { get; set; }
        //public string PPTSpeakerNotesLeft { get; set; }
        //public string PPTSpeakerNotesRight { get; set; }
        public string ShowPPTSpeakerNotesButtonNameLeft { get; set; }
        public string ShowPPTSpeakerNotesButtonNameRight { get; set; }
        public string ShowPPTSpeakerNotesButtonNameRightChanged { get; set; }
        private Visibility _showPPTSpeakerNotesButtonLeft;
        private Visibility _showPPTSpeakerNotesButtonRight;
        private Visibility _showPPTSpeakerNotesButtonRightChanged;

        private SolidColorBrush _showPPTButtonBackground;
        public SolidColorBrush ShowPPTNoteButtonBackground
        {
            get
            {
                return _showPPTButtonBackground;
            }

            set
            {
                _showPPTButtonBackground = value;
                OnPropertyChanged();
            }
        }

        public Visibility showPPTSpeakerNotesButtonLeft
        {
            get
            {
                return _showPPTSpeakerNotesButtonLeft;
            }

            set
            {
                _showPPTSpeakerNotesButtonLeft = value;
                OnPropertyChanged();
            }
        }

        public Visibility showPPTSpeakerNotesButtonRight
        {
            get
            {
                return _showPPTSpeakerNotesButtonRight;
            }

            set
            {
                _showPPTSpeakerNotesButtonRight = value;
                OnPropertyChanged();
            }
        }
        public Visibility showPPTSpeakerNotesButtonRightChanged
        {
            get
            {
                return _showPPTSpeakerNotesButtonRightChanged;
            }

            set
            {
                _showPPTSpeakerNotesButtonRightChanged = value;
                OnPropertyChanged();
            }
        }

        public Visibility PPTNoteGridLeftVisi
        {
            get
            {
                return _pptNoteGridLeftVisi;
            }

            set
            {
                _pptNoteGridLeftVisi = value;
                OnPropertyChanged();
            }
        }
        public Visibility PPTNoteGridRightVisi
        {
            get
            {
                return _pptNoteGridRightVisi;
            }

            set
            {
                _pptNoteGridRightVisi = value;
                OnPropertyChanged();
            }
        }

        public double BlurRadiusRight
        {
            get
            {
                return _blurRadiusRight;
            }

            set
            {
                _blurRadiusRight = value;
                OnPropertyChanged();
            }
        }
        public double BlurRadiusLeft
        {
            get
            {
                return _blurRadiusLeft;
            }

            set
            {
                _blurRadiusLeft = value;
                OnPropertyChanged();
            }
        }



        public Visibility ShowMask
        {
            get
            {
                return _showMask;
            }

            set
            {
                _showMask = value;
                OnPropertyChanged();
            }
        }

        public Visibility ShowHiddenLeft
        {
            get
            {
                return _showHiddenLeft;
            }

            set
            {
                _showHiddenLeft = value;
                OnPropertyChanged();
            }
        }
        public Visibility ShowHiddenRight
        {
            get
            {
                return _showHiddenRight;
            }

            set
            {
                _showHiddenRight = value;
                OnPropertyChanged();
            }
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
