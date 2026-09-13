using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ChuanHoa.Client.Core.Licensing;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    /// <summary>WPF account surface. Missing production services are explicit, never simulated.</summary>
    internal sealed class AccountSettingsWindow : Window
    {
        private readonly StackPanel _content;
        private readonly LocalAccessManager _access;
        private readonly Action _recognition;
        private readonly AccountAccessApiClient _api;

        public AccountSettingsWindow(LocalAccessManager access, Action recognition)
        {
            _access = access; _recognition = recognition; _api = new AccountAccessApiClient();
            Closed += (_, __) => _api.Dispose();
            Title = "Thiết lập · Chuẩn hóa";
            Width = 840; Height = 590; MinWidth = 720; MinHeight = 490;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            FontFamily = new FontFamily("Segoe UI"); FontSize = 14;
            Background = new SolidColorBrush(Color.FromRgb(243, 244, 246));
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(200) });
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            var navigation = new StackPanel { Margin = new Thickness(16, 24, 16, 16) };
            navigation.Children.Add(new TextBlock { Text = "CHUẨN HÓA", FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(12, 0, 0, 24) });
            AddNavigation(navigation, "Tài khoản", ShowAccount);
            AddNavigation(navigation, "Gói sử dụng", ShowPlans);
            AddNavigation(navigation, "Thiết bị", ShowDevices);
            AddNavigation(navigation, "Nhận diện văn bản", () => _recognition());
            grid.Children.Add(navigation);
            _content = new StackPanel { Margin = new Thickness(24) };
            var scroll = new ScrollViewer { Content = _content, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            Grid.SetColumn(scroll, 1); grid.Children.Add(scroll);
            Content = grid;
            ShowAccount();
        }

        private static void AddNavigation(Panel panel, string title, Action action)
        {
            var button = new Button { Content = title, Padding = new Thickness(12),
                Margin = new Thickness(0, 0, 0, 8), HorizontalContentAlignment = HorizontalAlignment.Left };
            button.Click += (_, __) => action(); panel.Children.Add(button);
        }

        private void Heading(string title, string description)
        {
            _content.Children.Clear();
            _content.Children.Add(new TextBlock { Text = title, FontSize = 28, FontWeight = FontWeights.SemiBold });
            _content.Children.Add(new TextBlock { Text = description, TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 10, 0, 24), Foreground = Brushes.DimGray });
        }

        private StackPanel Card(string title, string text)
        {
            var body = new StackPanel();
            body.Children.Add(new TextBlock { Text = title, FontSize = 17, FontWeight = FontWeights.SemiBold });
            body.Children.Add(new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 10, 0, 12) });
            _content.Children.Add(new Border { Background = Brushes.White, CornerRadius = new CornerRadius(8),
                BorderBrush = new SolidColorBrush(Color.FromRgb(220, 223, 228)), BorderThickness = new Thickness(1),
                Padding = new Thickness(20), Margin = new Thickness(0, 0, 0, 16), Child = body });
            return body;
        }

        private void ShowAccount()
        {
            Heading("Tài khoản", "Quản lý tài khoản và quyền sử dụng trên thiết bị này.");
            var card = Card("Đăng nhập tài khoản", _api.IsConfigured ? "Tài khoản được bind với một thiết bị tại một thời điểm; phiên đăng nhập có hiệu lực tối đa 72 giờ." : "Nhập thông tin để đăng nhập khi API Chuẩn Hóa đã được cấu hình HTTPS.");
            var email = new TextBox { MinWidth = 360, Margin = new Thickness(0, 4, 0, 8) };
            var password = new PasswordBox { MinWidth = 360, Margin = new Thickness(0, 4, 0, 8) };
            card.Children.Add(new TextBlock { Text = "Email" }); card.Children.Add(email);
            card.Children.Add(new TextBlock { Text = "Mật khẩu", Margin = new Thickness(0, 8, 0, 0) }); card.Children.Add(password);
            var login = new Button { Content = "Đăng nhập", Padding = new Thickness(16, 8, 16, 8), HorizontalAlignment = HorizontalAlignment.Left };
            login.Click += async (_, __) =>
            {
                login.IsEnabled = false;
                try
                {
                    var session = await _api.LoginAsync(email.Text, password.Password, _access.DeviceThumbprint);
                    _access.SaveSessionToken(session.SessionToken);
                    _access.SetAccessMode(new AccessModeState(AccessMode.Account, session.IssuedAtUtc, session.ExpiresAtUtc, session.DeviceThumbprint));
                    MessageBox.Show("Đăng nhập thành công.", "Chuẩn hóa", MessageBoxButton.OK, MessageBoxImage.Information); ShowAccount();
                }
                catch (Exception error) { MessageBox.Show(error.Message, "Không thể đăng nhập", MessageBoxButton.OK, MessageBoxImage.Warning); }
                finally { login.IsEnabled = true; }
            };
            card.Children.Add(login);
            var register = new Button { Content = "Đăng ký tài khoản", Padding = new Thickness(16, 8, 16, 8), Margin = new Thickness(0, 8, 0, 0), HorizontalAlignment = HorizontalAlignment.Left };
            register.Click += async (_, __) =>
            {
                register.IsEnabled = false;
                try
                {
                    var session = await _api.RegisterAsync(email.Text, password.Password, email.Text, _access.DeviceThumbprint);
                    _access.SaveSessionToken(session.SessionToken);
                    _access.SetAccessMode(new AccessModeState(AccessMode.Account, session.IssuedAtUtc, session.ExpiresAtUtc, session.DeviceThumbprint));
                    MessageBox.Show("Đăng ký thành công.", "Chuẩn hóa", MessageBoxButton.OK, MessageBoxImage.Information); ShowAccount();
                }
                catch (Exception error) { MessageBox.Show(error.Message, "Không thể đăng ký", MessageBoxButton.OK, MessageBoxImage.Warning); }
                finally { register.IsEnabled = true; }
            };
            card.Children.Add(register);
            if (_access.AccessModeState?.Mode == AccessMode.Account)
            {
                var logout = new Button { Content = "Đăng xuất", Padding = new Thickness(16, 8, 16, 8), Margin = new Thickness(8, 0, 0, 0), HorizontalAlignment = HorizontalAlignment.Left };
                logout.Click += async (_, __) =>
                {
                    logout.IsEnabled = false;
                    try { await _api.LogoutAsync(_access.LoadSessionToken() ?? string.Empty); _access.ClearSessionToken(); _access.SetAccessMode(null); ShowAccount(); }
                    catch (Exception error) { MessageBox.Show(error.Message, "Không thể đăng xuất", MessageBoxButton.OK, MessageBoxImage.Warning); logout.IsEnabled = true; }
                };
                card.Children.Add(logout);
            }
            var activation = Card("Kích hoạt VIP", "VIP key được bind theo thiết bị và không yêu cầu đăng nhập tài khoản.");
            var key = new TextBox { MinWidth = 360, Margin = new Thickness(0, 4, 0, 8) };
            activation.Children.Add(new TextBlock { Text = "Mã kích hoạt" }); activation.Children.Add(key);
            var activate = new Button { Content = "Kích hoạt", Padding = new Thickness(16, 8, 16, 8), HorizontalAlignment = HorizontalAlignment.Left };
            activate.Click += async (_, __) =>
            {
                activate.IsEnabled = false;
                try
                {
                    var result = await _api.ActivateAsync(key.Text, _access.DeviceThumbprint);
                    // For non-expiring keys the signed lease remains the authoritative expiry;
                    // the activation mode itself is still device-bound and never grants features.
                    var expiry = result.ExpiresAtUtc ?? DateTimeOffset.MaxValue;
                    _access.SetAccessMode(new AccessModeState(AccessMode.ActivationKey, DateTimeOffset.UtcNow, expiry, _access.DeviceThumbprint));
                    MessageBox.Show("Đã kích hoạt VIP.", "Chuẩn hóa", MessageBoxButton.OK, MessageBoxImage.Information); ShowAccount();
                }
                catch (Exception error) { MessageBox.Show(error.Message, "Không thể kích hoạt", MessageBoxButton.OK, MessageBoxImage.Warning); }
                finally { activate.IsEnabled = true; }
            };
            activation.Children.Add(activate);
            Card("Giấy phép cục bộ", _access.DescribeStatus());
        }

        private void ShowPlans()
        {
            Heading("Gói sử dụng", "Thanh toán và nâng cấp ngay trong tiện ích sau khi kết nối tài khoản.");
            Card("Miễn phí", "• Kiểm tra thể thức\n• Kiểm tra chính tả\n• Tất cả chức năng Bảng biểu và hình ảnh");
            var card = Card("Gói trả phí duy nhất", "Mở khóa các chức năng ngoài gói miễn phí. Giá và thời hạn của gói lấy từ máy chủ; các thay đổi giá là phiên bản giá của cùng một gói. Tài khoản và thanh toán payOS chưa cấu hình hoàn chỉnh; chưa tạo đơn hoặc thu tiền.");
            card.Children.Add(new Button { Content = "Mua gói qua payOS", IsEnabled = false,
                Padding = new Thickness(16, 8, 16, 8), HorizontalAlignment = HorizontalAlignment.Left,
                ToolTip = "Cần kết nối tài khoản và dịch vụ thanh toán đã cấu hình." });
        }

        private void ShowDevices()
        {
            Heading("Thiết bị", "Danh sách thiết bị và trạng thái đăng ký lấy từ tài khoản trên máy chủ.");
            Card("Máy hiện tại", Environment.MachineName);
            Card("Chưa kết nối", "Không thể liệt kê hoặc thu hồi thiết bị trước khi đăng nhập. Không thay đổi giấy phép cục bộ khi dịch vụ chưa sẵn sàng.");
        }
    }
}
