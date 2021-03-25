using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Xml;

namespace DocCompareWPF.UIhelper
{
    public static class TextBlockHelper
    {
        #region FormattedText Attached dependency property

        public static string GetFormattedText(DependencyObject obj)
        {
            return (string)obj.GetValue(FormattedTextProperty);
        }

        public static void SetFormattedText(DependencyObject obj, string value)
        {
            obj.SetValue(FormattedTextProperty, value);
        }

        public static readonly DependencyProperty FormattedTextProperty =
            DependencyProperty.RegisterAttached("FormattedText",
            typeof(string),
            typeof(TextBlockHelper),
            new UIPropertyMetadata("", FormattedTextChanged));

        private static void FormattedTextChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            string value = e.NewValue as string;

            TextBlock textBlock = sender as TextBlock;

            if (textBlock != null)
            {
                textBlock.Inlines.Clear();
                textBlock.Inlines.Add(Process(value));
            }
        }

        #endregion

        static Inline Process(string value)
        {
            XmlDocument doc = new XmlDocument();

            if (value != null)
                doc.LoadXml(value);

            Span span = new Span();

            if (doc.ChildNodes.Count != 0)
                InternalProcess(span, doc.ChildNodes[1]);

            return span;
        }

        private static void InternalProcess(Span span, XmlNode xmlNode)
        {
            foreach (XmlNode child in xmlNode)
            {
                if (child is XmlText)
                {
                    span.Inlines.Add(new Run(child.InnerText));
                }
                else if (child is XmlElement)
                {
                    Span spanItem = new Span();
                    InternalProcess(spanItem, child);
                    switch (child.Name.ToUpper())
                    {
                        case "B":
                        case "BOLD":
                            Bold bold = new Bold(spanItem);
                            span.Inlines.Add(bold);
                            break;
                        case "I":
                        case "ITALIC":
                            Italic italic = new Italic(spanItem);
                            span.Inlines.Add(italic);
                            break;
                        case "U":
                        case "UNDERLINE":
                            Underline underline = new Underline(spanItem);
                            span.Inlines.Add(underline);
                            break;
                        case "D":
                        case "DELETE":
                            spanItem.Background = new SolidColorBrush(Color.FromArgb(128, 255, 44, 108));
                            spanItem.Foreground = Brushes.Transparent;
                            span.Inlines.Add(spanItem);
                            break;
                        case "IN":
                        case "INSERT":
                            spanItem.Background = new SolidColorBrush(Color.FromArgb(128, 255, 44, 108));
                            span.Inlines.Add(spanItem);
                            break;
                    }
                }
            }
        }
    }

}
