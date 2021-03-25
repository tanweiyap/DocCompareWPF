using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Effects;

namespace DocCompareWPF.UIhelper
{
    public class SideGridItemRight : INotifyPropertyChanged
    {
        private Color _color;
        private Effect _effect;
        private Visibility _showMask;

        public event PropertyChangedEventHandler PropertyChanged;

        private Visibility _showHidden;
        private Visibility _forceAlignButtonVisi;
        private Visibility _forceAlignButtonInvalidVisi;

        private Visibility _diffVisi;
        private Visibility _noDiffVisi;

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
        public Visibility DiffVisi
        {
            get
            {
                return _diffVisi;
            }

            set
            {
                _diffVisi = value;
                OnPropertyChanged();
            }
        }

        public Visibility NoDiffVisi
        {
            get
            {
                return _noDiffVisi;
            }

            set
            {
                _noDiffVisi = value;
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
        public string ImgMaskName { get; set; }
        public string ImgName { get; set; }
        public Thickness Margin { get; set; }

        private string _pathToImg;
        public string PathToImg { get { return _pathToImg; } set { _pathToImg = value; OnPropertyChanged(); } }

        private string _pathToImgDummy;
        public string PathToImgDummy { get { return _pathToImgDummy; } set { _pathToImgDummy = value; OnPropertyChanged(); } }

        private string _pathToMask;
        public string PathToMask { get { return _pathToMask; } set { _pathToMask = value; OnPropertyChanged(); } }
        public bool RemoveForceAlignButtonEnable { get; set; }
        public string RemoveForceAlignButtonName { get; set; }

        public Visibility RemoveForceAlignButtonVisibility { get; set; }

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

        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
