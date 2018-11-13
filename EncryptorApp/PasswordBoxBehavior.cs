using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace EncryptorApp
{
    /// <summary>
    ///     Simple behavior for the PasswordBox to provide a watermark text element.
    /// </summary>
    public static class PasswordBoxBehavior
    {
        private const string WatermarkId = "_pboxWatermark";

        /// <summary>
        ///     Backing storage key for the text property
        /// </summary>
        public static readonly DependencyProperty WatermarkProperty =
            DependencyProperty.RegisterAttached("Watermark", typeof(string), typeof(PasswordBoxBehavior),
                                                new PropertyMetadata("", OnWatermarkChanged));

        /// <summary>
        ///     Gets the watermark text
        /// </summary>
        /// <param name="pbox"></param>
        /// <returns></returns>
        public static string GetWatermark(PasswordBox pbox)
        {
            return (string)pbox.GetValue(WatermarkProperty);
        }

        /// <summary>
        ///     Sets the watermark text
        /// </summary>
        /// <param name="pbox"></param>
        /// <param name="text"></param>
        public static void SetWatermark(PasswordBox pbox, string text)
        {
            pbox.SetValue(WatermarkProperty, text);
        }

        /// <summary>
        ///     Called when the watermark is changed.
        /// </summary>
        /// <param name="dpo"></param>
        /// <param name="e"></param>
        private static void OnWatermarkChanged(DependencyObject dpo, DependencyPropertyChangedEventArgs e)
        {
            var pbox = dpo as PasswordBox;
            if (pbox == null)
                return;

            pbox.PasswordChanged += PboxOnPasswordChanged;
            pbox.GotFocus += PboxOnGotFocus;
            pbox.LostFocus += PboxOnLostFocus;
            pbox.Loaded += PboxOnLoaded;

            string text = (e.NewValue ?? "").ToString();
            if (string.IsNullOrEmpty(text))
            {
                RemoveWatermarkElement(pbox);
            }
            else
            {
                AddWatermarkElement(pbox, text);
            }
        }

        /// <summary>
        ///     Called when the PasswordBox is loaded.  This adds the watermark if one is present.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="routedEventArgs"></param>
        private static void PboxOnLoaded(object sender, RoutedEventArgs routedEventArgs)
        {
            var pbox = (PasswordBox)sender;
            string text = GetWatermark(pbox);
            if (string.IsNullOrEmpty(text))
            {
                RemoveWatermarkElement(pbox);
            }
            else
            {
                AddWatermarkElement(pbox, text);
            }
        }

        /// <summary>
        ///     Called when the PasswordBox loses focus - this adds the watermark if necessary.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="routedEventArgs"></param>
        private static void PboxOnLostFocus(object sender, RoutedEventArgs routedEventArgs)
        {
            var pbox = (PasswordBox)sender;
            if (pbox.Password.Length == 0)
            {
                AddWatermarkElement(pbox, GetWatermark(pbox));
            }
        }

        /// <summary>
        ///     Called when the PasswordBox gets focus - this removes any watermark.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="routedEventArgs"></param>
        private static void PboxOnGotFocus(object sender, RoutedEventArgs routedEventArgs)
        {
            var pbox = (PasswordBox)sender;
            RemoveWatermarkElement(pbox);
        }

        /// <summary>
        ///     This is called when the password is changed in the PasswordBox and removes the watermark.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="routedEventArgs"></param>
        private static void PboxOnPasswordChanged(object sender, RoutedEventArgs routedEventArgs)
        {
            var pbox = (PasswordBox)sender;
            if (pbox.Password.Length > 0)
            {
                RemoveWatermarkElement(pbox);
            }
        }

        /// <summary>
        ///     Simple method to add a new TextBlock into the visual tree of the
        ///     PasswordBox which will present the watermark.
        /// </summary>
        /// <param name="pbox"></param>
        /// <param name="text"></param>
        private static void AddWatermarkElement(PasswordBox pbox, string text)
        {
            var wmTb = pbox.FindVisualChildByName<TextBlock>(WatermarkId);
            if (wmTb == null)
            {
                var fe = pbox.FindVisualChildByName<ScrollViewer>("ContentElement");
                if (fe != null)
                {
                    var panelOwner = fe.FindVisualParent<Panel>();
                    if (panelOwner != null)
                    {
                        // Add the TextBlock.
                        var tb = new TextBlock
                        {
                            Name = WatermarkId,
                            Text = text,
                            HorizontalAlignment = HorizontalAlignment.Left,
                            VerticalAlignment = VerticalAlignment.Center,
                            Margin = new Thickness(3, 0, 0, 0),
                            Foreground = new SolidColorBrush(Colors.Gray)
                        };
                        int index = panelOwner.Children.IndexOf(fe);
                        panelOwner.Children.Insert(index + 1, tb);
                    }
                }
            }
        }

        /// <summary>
        ///     Simple method to remove the TextBlock from the PasswordBox
        ///     visual tree.
        /// </summary>
        /// <param name="pbox"></param>
        private static void RemoveWatermarkElement(PasswordBox pbox)
        {
            var wmTb = pbox.FindVisualChildByName<TextBlock>(WatermarkId);
            if (wmTb != null)
            {
                var panelOwner = wmTb.FindVisualParent<Panel>();
                if (panelOwner != null)
                {
                    panelOwner.Children.Remove(wmTb);
                }
            }
        }
    }
}
