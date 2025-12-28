using System.Windows;
using System.Windows.Input;

namespace ExcelFileLocator
{
    public partial class CustomMessageBox : Window
    {
        public MessageBoxResult Result { get; private set; } = MessageBoxResult.None;

        private CustomMessageBox(string message, string title, MessageBoxButton button, MessageBoxImage icon, MessageBoxResult defaultResult)
        {
            InitializeComponent();

            txtTitle.Text = title;
            txtMessage.Text = message;

            // 设置图标和颜色
            SetIcon(icon);

            // 设置按钮
            SetButtons(button, defaultResult);

            // 激活窗口
            this.Activated += (s, e) => this.Focus();
        }

        private void SetIcon(MessageBoxImage icon)
        {
            switch (icon)
            {
                case MessageBoxImage.Information:
                    txtIcon.Text = "ℹ";
                    txtIcon.Foreground = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#4A90E2"));
                    break;
                case MessageBoxImage.Question:
                    txtIcon.Text = "?";
                    txtIcon.Foreground = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#4A90E2"));
                    break;
                case MessageBoxImage.Warning:
                    txtIcon.Text = "⚠";
                    txtIcon.Foreground = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F39C12"));
                    break;
                case MessageBoxImage.Error:
                    txtIcon.Text = "✖";
                    txtIcon.Foreground = new System.Windows.Media.SolidColorBrush(
                        (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E74C3C"));
                    break;
                default:
                    txtIcon.Visibility = Visibility.Collapsed;
                    break;
            }
        }

        private void SetButtons(MessageBoxButton button, MessageBoxResult defaultResult)
        {
            switch (button)
            {
                case MessageBoxButton.OK:
                    btnOK.Visibility = Visibility.Visible;
                    btnOK.IsDefault = true;
                    break;

                case MessageBoxButton.OKCancel:
                    btnOK.Visibility = Visibility.Visible;
                    btnCancel.Visibility = Visibility.Visible;
                    btnOK.IsDefault = (defaultResult == MessageBoxResult.OK || defaultResult == MessageBoxResult.None);
                    btnCancel.IsDefault = (defaultResult == MessageBoxResult.Cancel);
                    break;

                case MessageBoxButton.YesNo:
                    btnYes.Visibility = Visibility.Visible;
                    btnNo.Visibility = Visibility.Visible;
                    btnYes.IsDefault = (defaultResult == MessageBoxResult.Yes || defaultResult == MessageBoxResult.None);
                    btnNo.IsDefault = (defaultResult == MessageBoxResult.No);
                    break;

                case MessageBoxButton.YesNoCancel:
                    btnYes.Visibility = Visibility.Visible;
                    btnNo.Visibility = Visibility.Visible;
                    btnCancel.Visibility = Visibility.Visible;
                    btnYes.IsDefault = (defaultResult == MessageBoxResult.Yes || defaultResult == MessageBoxResult.None);
                    btnNo.IsDefault = (defaultResult == MessageBoxResult.No);
                    btnCancel.IsDefault = (defaultResult == MessageBoxResult.Cancel);
                    break;
            }
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            Result = MessageBoxResult.OK;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            Result = MessageBoxResult.Cancel;
            this.Close();
        }

        private void BtnYes_Click(object sender, RoutedEventArgs e)
        {
            Result = MessageBoxResult.Yes;
            this.Close();
        }

        private void BtnNo_Click(object sender, RoutedEventArgs e)
        {
            Result = MessageBoxResult.No;
            this.Close();
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        // 静态方法，用于显示消息框
        public static MessageBoxResult Show(string message, string title = "提示",
            MessageBoxButton button = MessageBoxButton.OK,
            MessageBoxImage icon = MessageBoxImage.Information,
            MessageBoxResult defaultResult = MessageBoxResult.None)
        {
            var messageBox = new CustomMessageBox(message, title, button, icon, defaultResult);
            messageBox.ShowDialog();
            return messageBox.Result;
        }

        public static MessageBoxResult Show(Window owner, string message, string title = "提示",
            MessageBoxButton button = MessageBoxButton.OK,
            MessageBoxImage icon = MessageBoxImage.Information,
            MessageBoxResult defaultResult = MessageBoxResult.None)
        {
            var messageBox = new CustomMessageBox(message, title, button, icon, defaultResult)
            {
                Owner = owner
            };
            messageBox.ShowDialog();
            return messageBox.Result;
        }
    }
}