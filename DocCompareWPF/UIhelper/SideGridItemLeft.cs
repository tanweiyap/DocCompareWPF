using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Effects;

namespace DocCompareWPF.UIhelper
{
    public class SideGridItemLeft : INotifyPropertyChanged
    {
        private Color _color;
        private Effect _effect;

        public event PropertyChangedEventHandler PropertyChanged;

        private Visibility _forceAlignButtonVisi;
        private Visibility _forceAlignButtonInvalidVisi;

        private Visibility _showHidden;

        public Visibility ShowHidden
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
        public Visibility ForceAlignButtonVisi
        {
            get
            {
                return _forceAlignButtonVisi;
            }

            set
            {
                _forceAlignButtonVisi = value;
                OnPropertyChanged();
            }
        }
        public Visibility ForceAlignButtonInvalidVisi
        {
            get
            {
                return _forceAlignButtonInvalidVisi;
            }

            set
            {
                _forceAlignButtonInvalidVisi = value;
                OnPropertyChanged();
            }
        }

        public Color BackgroundBrush
        {
            get
            {
                return _color;
            }

            set
            {
                _color = value;
                OnPropertyChanged();
            }
        }

        public string ForceAlignButtonName { get; set; }
        public string ForceAlignInvalidButtonName { get; set; }

        public Effect GridEffect
        {
            get
            {
                return _effect;
            }

            set
            {
                _effect = value;
                OnPropertyChanged();
            }
        }

        public string GridName { get; set; }
        public string ImgDummyName { get; set; }
        public string ImgGridName { get; set; }
        public string ImgName { get; set; }
        public Thickness Margin { get; set; }
        public string PageNumberLabel { get; set; }

        private string _pathToImg;
        public string PathToImg { get { return _pathToImg; } set { _pathToImg = value; OnPropertyChanged(); } }

        private string _pathToImgDummy;
        public string PathToImgDummy { get { return _pathToImgDummy; } set { _pathToImgDummy = value; OnPropertyChanged(); } }
        public string RemoveForceAlignButtonName { get; set; }
        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
