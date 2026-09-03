using System;
using System.Drawing;
using System.Windows.Forms;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    internal sealed class QrCodeInputDialog : Form
    {
        private readonly TextBox _content;

        private QrCodeInputDialog()
        {
            Text = "Chèn mã QR";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(520, 250);
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            var label = new Label
            {
                AutoSize = true,
                Location = new Point(16, 16),
                Text = "Nhập nội dung cần mã hóa (tối đa 800 ký tự):"
            };
            _content = new TextBox
            {
                Location = new Point(16, 42),
                Size = new Size(488, 150),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                MaxLength = 800,
                AcceptsReturn = true,
                AcceptsTab = false
            };
            var cancel = new Button
            {
                Text = "Hủy",
                DialogResult = DialogResult.Cancel,
                Location = new Point(328, 207),
                Size = new Size(80, 30)
            };
            var confirm = new Button
            {
                Text = "Chèn QR",
                DialogResult = DialogResult.OK,
                Location = new Point(424, 207),
                Size = new Size(80, 30)
            };
            Controls.Add(label);
            Controls.Add(_content);
            Controls.Add(cancel);
            Controls.Add(confirm);
            AcceptButton = confirm;
            CancelButton = cancel;
        }

        public static string? Prompt(IWin32Window? owner)
        {
            using (var dialog = new QrCodeInputDialog())
            {
                if (dialog.ShowDialog(owner) != DialogResult.OK) return null;
                var value = dialog._content.Text.Trim();
                if (value.Length == 0)
                {
                    MessageBox.Show(dialog, "Nội dung mã QR không được để trống.", "Chèn mã QR",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }
                return value;
            }
        }
    }
}
