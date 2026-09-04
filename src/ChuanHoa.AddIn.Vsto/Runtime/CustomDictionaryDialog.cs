using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ChuanHoa.Client.Core.Lexicon;

namespace ChuanHoa.AddIn.Vsto.Runtime
{
    internal sealed class CustomDictionaryDialog : Form
    {
        private readonly ListBox _wordListBox;
        private readonly TextBox _searchBox;
        private readonly TextBox _newWordBox;
        private readonly Label _countLabel;

        private CustomDictionaryDialog()
        {
            Text = "Quản lý Từ điển cá nhân & Danh sách bỏ qua";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(540, 440);
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            var titleLabel = new Label
            {
                AutoSize = true,
                Location = new Point(16, 14),
                Text = "Từ vựng riêng, tên viết tắt hoặc thuật ngữ chuyên ngành (không bị báo lỗi):"
            };

            var searchLabel = new Label
            {
                AutoSize = true,
                Location = new Point(16, 40),
                Text = "Tìm kiếm:"
            };

            _searchBox = new TextBox
            {
                Location = new Point(80, 37),
                Size = new Size(444, 25)
            };
            _searchBox.TextChanged += (s, e) => RefreshWordList();

            _wordListBox = new ListBox
            {
                Location = new Point(16, 68),
                Size = new Size(390, 260),
                IntegralHeight = false
            };

            _countLabel = new Label
            {
                AutoSize = true,
                Location = new Point(16, 334),
                Text = "Tổng số từ: 0"
            };

            var btnDelete = new Button
            {
                Text = "Xóa từ",
                Location = new Point(416, 68),
                Size = new Size(108, 30)
            };
            btnDelete.Click += (s, e) => DeleteSelectedWord();

            var btnClearIgnores = new Button
            {
                Text = "Xóa bỏ qua...",
                Location = new Point(416, 106),
                Size = new Size(108, 30)
            };
            btnClearIgnores.Click += (s, e) => ClearDocumentIgnores();

            var newWordLabel = new Label
            {
                AutoSize = true,
                Location = new Point(16, 360),
                Text = "Thêm từ mới:"
            };

            _newWordBox = new TextBox
            {
                Location = new Point(105, 357),
                Size = new Size(301, 25)
            };
            _newWordBox.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    AddNewWord();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            };

            var btnAdd = new Button
            {
                Text = "Thêm",
                Location = new Point(416, 355),
                Size = new Size(108, 28)
            };
            btnAdd.Click += (s, e) => AddNewWord();

            var btnClose = new Button
            {
                Text = "Đóng",
                DialogResult = DialogResult.OK,
                Location = new Point(416, 396),
                Size = new Size(108, 32)
            };

            Controls.Add(titleLabel);
            Controls.Add(searchLabel);
            Controls.Add(_searchBox);
            Controls.Add(_wordListBox);
            Controls.Add(_countLabel);
            Controls.Add(btnDelete);
            Controls.Add(btnClearIgnores);
            Controls.Add(newWordLabel);
            Controls.Add(_newWordBox);
            Controls.Add(btnAdd);
            Controls.Add(btnClose);

            CancelButton = btnClose;

            RefreshWordList();
        }

        private void RefreshWordList()
        {
            var filter = _searchBox.Text.Trim();
            var allWords = PersonalDictionaryManager.Instance.GetUserWords();
            _wordListBox.BeginUpdate();
            _wordListBox.Items.Clear();

            var filtered = string.IsNullOrEmpty(filter)
                ? allWords
                : allWords.Where(w => w.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0);

            foreach (var word in filtered)
            {
                _wordListBox.Items.Add(word);
            }
            _wordListBox.EndUpdate();
            _countLabel.Text = "Tổng số từ: " + allWords.Count.ToString();
        }

        private void AddNewWord()
        {
            var word = _newWordBox.Text.Trim();
            if (string.IsNullOrEmpty(word))
            {
                return;
            }

            PersonalDictionaryManager.Instance.AddUserWord(word);
            _newWordBox.Clear();
            RefreshWordList();
            _wordListBox.SelectedItem = word;
        }

        private void DeleteSelectedWord()
        {
            if (_wordListBox.SelectedItem == null)
            {
                MessageBox.Show(this, "Hãy chọn một từ trong danh sách để xóa.", "Từ điển cá nhân",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var word = _wordListBox.SelectedItem.ToString();
            if (!string.IsNullOrEmpty(word))
            {
                PersonalDictionaryManager.Instance.RemoveUserWord(word);
                RefreshWordList();
            }
        }

        private void ClearDocumentIgnores()
        {
            var result = MessageBox.Show(this,
                "Bạn có muốn đặt lại danh sách các từ đã tạm bỏ qua trong các tài liệu hiện tại không?",
                "Danh sách bỏ qua",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // Clear any document session ignores
                PersonalDictionaryManager.Instance.ClearDocumentIgnores(string.Empty);
                MessageBox.Show(this, "Đã xóa toàn bộ danh sách từ bỏ qua tạm thời.", "Từ điển cá nhân",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public static void Prompt(IWin32Window? owner = null)
        {
            using (var dialog = new CustomDictionaryDialog())
            {
                ((Form)dialog).ShowDialog(owner);
            }
        }
    }
}
