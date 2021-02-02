using System;
using System.Windows;
using System.Windows.Controls;

namespace DocCompareWPF
{
    /// <summary>
    /// Interaktionslogik für DialogBubble.xaml
    /// </summary>
    public partial class DialogBubble : UserControl
    {
        public String Message
        {
            get { return (String)GetValue(LabelProperty); }
            set { SetValue(LabelProperty, value); }
        }

        public static readonly DependencyProperty LabelProperty =
            DependencyProperty.Register("Message", typeof(string),
              typeof(DialogBubble), new PropertyMetadata(""));

        public Point StartPoint
        {
            get { return (Point)GetValue(StartPointProperty); }
            set { SetValue(StartPointProperty, value); }
        }

        public static readonly DependencyProperty StartPointProperty =
            DependencyProperty.Register("StartPoint", typeof(Point),
              typeof(DialogBubble), new PropertyMetadata(new Point(0, 0)));

        public Point TopPoint
        {
            get { return (Point)GetValue(TopPointProperty); }
            set { SetValue(TopPointProperty, value); }
        }

        public static readonly DependencyProperty TopPointProperty =
            DependencyProperty.Register("TopPoint", typeof(Point),
              typeof(DialogBubble), new PropertyMetadata(new Point(0, 0)));

        public Point EndPoint
        {
            get { return (Point)GetValue(EndPointProperty); }
            set { SetValue(EndPointProperty, value); }
        }

        public static readonly DependencyProperty EndPointProperty =
            DependencyProperty.Register("EndPoint", typeof(Point),
              typeof(DialogBubble), new PropertyMetadata(new Point(0, 0)));

        public Thickness RectInnerMargin
        {
            get { return (Thickness)GetValue(RectInnerMarginProperty); }
            set { SetValue(RectInnerMarginProperty, value); }
        }

        public static readonly DependencyProperty RectInnerMarginProperty =
            DependencyProperty.Register("RectInnerMargin", typeof(Thickness),
              typeof(DialogBubble), new PropertyMetadata(new Thickness(20, 60, 20, 20)));

        public Point RectPoint1
        {
            get { return (Point)GetValue(RectPoint1Property); }
            set { SetValue(RectPoint1Property, value); }
        }

        public static readonly DependencyProperty RectPoint1Property =
            DependencyProperty.Register("RectPoint1", typeof(Point),
              typeof(DialogBubble), new PropertyMetadata(new Point(0, 0)));

        public Point RectPoint2
        {
            get { return (Point)GetValue(RectPoint2Property); }
            set { SetValue(RectPoint2Property, value); }
        }

        public static readonly DependencyProperty RectPoint2Property =
            DependencyProperty.Register("RectPoint2", typeof(Point),
              typeof(DialogBubble), new PropertyMetadata(new Point(0, 0)));

        public Point RectPoint3
        {
            get { return (Point)GetValue(RectPoint3Property); }
            set { SetValue(RectPoint3Property, value); }
        }

        public static readonly DependencyProperty RectPoint3Property =
            DependencyProperty.Register("RectPoint3", typeof(Point),
              typeof(DialogBubble), new PropertyMetadata(new Point(0, 0)));

        public Point RectPoint4
        {
            get { return (Point)GetValue(RectPoint4Property); }
            set { SetValue(RectPoint4Property, value); }
        }

        public static readonly DependencyProperty RectPoint4Property =
            DependencyProperty.Register("RectPoint4", typeof(Point),
              typeof(DialogBubble), new PropertyMetadata(new Point(0, 0)));

        /*
        public ArrowPosition ArrPosition
        {
            get { return (ArrowPosition)GetValue(ArrPositionProperty); }
            set
            {
                SetValue(ArrPositionProperty, value);
            }
        }

        public static readonly DependencyProperty ArrPositionProperty =
            DependencyProperty.Register("ArrPosition", typeof(ArrowPosition),
              typeof(DialogBubble), new PropertyMetadata(new PropertyChangedCallback(OnEnumArrPositionChanged)));
        public enum ArrowPosition
        {
            TOPLEFT,
            TOPMIDDLE,
            TOPRIGHT,
            BOTTOMLEFT,
            BOTTOMMIDDLE,
            BOTTOMRIGHT
        }

        private static void OnEnumArrPositionChanged(DependencyObject control, DependencyPropertyChangedEventArgs args)
        {
            switch ((ArrowPosition) args.NewValue)
            {
                case ArrowPosition.TOPLEFT:
                    ((DialogBubble)control).ArrPathFigure.StartPoint = new Point(20, 40);
                    ((DialogBubble)control).ArrPathTopPoint.Point = new Point(50, 40);
                    ((DialogBubble)control).ArrPathEndPoint.Point = new Point(80, 40);
                    ((DialogBubble)control).RectMargin = new Thickness(20, 60, 20, 20);
                    break;

                case ArrowPosition.TOPMIDDLE:
                    ((DialogBubble)control).StartPoint = new Point(120, 40);
                    ((DialogBubble)control).TopPoint = new Point(150, 40);
                    ((DialogBubble)control).EndPoint = new Point(180, 40);
                    ((DialogBubble)control).RectMargin = new Thickness(20, 60, 20, 20);
                    break;

                case ArrowPosition.TOPRIGHT:
                    ((DialogBubble)control).StartPoint = new Point(210, 40);
                    ((DialogBubble)control).TopPoint = new Point(240, 40);
                    ((DialogBubble)control).EndPoint = new Point(270, 40);
                    ((DialogBubble)control).RectMargin = new Thickness(20, 60, 20, 20);
                    break;
            }
        }
        */

        public DialogBubble()
        {
            InitializeComponent();
            DataContext = this;
        }
    }
}