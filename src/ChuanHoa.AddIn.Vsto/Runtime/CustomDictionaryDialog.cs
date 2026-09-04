using System;
using System.Drawing;
using System.Linq;
using System.Text;
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
        private readonly Label _statusLabel;
        private readonly Button _deleteButton;
        private readonly string? _currentDocumentId;
        private readonly string? _selectedText;

        private CustomDictionaryDialog(string? currentDocumentId, string? selectedText)
        {
            _currentDocumentId = currentDocumentId;
            _selectedText = NormalizeOptionalEntry(selectedText);
            Text = "Quản lý Từ điển cá nhân & Danh sách bỏ qua";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
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
                Name = "txtDictionarySearch",
                AccessibleName = "Tìm kiếm trong từ điển cá nhân",
                Location = new Point(80, 37),
                Size = new Size(444, 25),
                TabIndex = 0
            };
            _searchBox.TextChanged += (s, e) => RefreshWordList();

            _wordListBox = new ListBox
            {
                Name = "lstDictionaryWords",
                AccessibleName = "Danh sách từ điển cá nhân",
                Location = new Point(16, 68),
                Size = new Size(390, 260),
                IntegralHeight = false,
                TabIndex = 1
            };

            _countLabel = new Label
            {
                AutoSize = true,
                Location = new Point(16, 334),
                Text = "Tổng số từ: 0"
            };

            _deleteButton = new Button
            {
                Name = "btnDeleteDictionaryWord",
                Text = "Xóa từ",
                Location = new Point(416, 68),
                Size = new Size(108, 30),
                Enabled = false,
                TabIndex = 2
            };
            _deleteButton.Click += (s, e) => DeleteSelectedWord();
            _wordListBox.SelectedIndexChanged += (s, e) =>
                _deleteButton.Enabled = _wordListBox.SelectedItem != null;

            var btnClearIgnores = new Button
            {
                Name = "btnClearDocumentIgnores",
                Text = "Xóa bỏ qua...",
                Location = new Point(416, 106),
                Size = new Size(108, 30),
                TabIndex = 3
            };
            btnClearIgnores.Click += (s, e) => ClearDocumentIgnores();

            var btnAddSelection = new Button
            {
                Name = "btnAddSelectedText",
                Text = "Thêm phần chọn",
                Location = new Point(416, 144),
                Size = new Size(108, 38),
                Enabled = !string.IsNullOrWhiteSpace(_selectedText),
                TabIndex = 4
            };
            btnAddSelection.Click += (s, e) => AddSelectedText();

            var btnIgnoreSelection = new Button
            {
                Name = "btnIgnoreSelectedText",
                Text = "Bỏ qua tài liệu",
                Location = new Point(416, 190),
                Size = new Size(108, 38),
                Enabled = !string.IsNullOrWhiteSpace(_selectedText) &&
                    !string.IsNullOrWhiteSpace(_currentDocumentId),
                TabIndex = 5
            };
            btnIgnoreSelection.Click += (s, e) => IgnoreSelectedText();

            var newWordLabel = new Label
            {
                AutoSize = true,
                Location = new Point(16, 360),
                Text = "Thêm từ mới:"
            };

            _newWordBox = new TextBox
            {
                Name = "txtNewDictionaryWord",
                AccessibleName = "Từ hoặc cụm từ mới",
                Location = new Point(105, 357),
                Size = new Size(301, 25),
                TabIndex = 6
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
                Name = "btnAddDictionaryWord",
                Text = "Thêm",
                Location = new Point(416, 355),
                Size = new Size(108, 28),
                TabIndex = 7
            };
            btnAdd.Click += (s, e) => AddNewWord();

            var btnClose = new Button
            {
                Name = "btnCloseDictionary",
                Text = "Đóng",
                DialogResult = DialogResult.OK,
                Location = new Point(416, 396),
                Size = new Size(108, 32),
                TabIndex = 8
            };

            _statusLabel = new Label
            {
                AutoSize = false,
                Location = new Point(16, 390),
                Size = new Size(390, 38),
                ForeColor = Color.Firebrick
            };

            Controls.Add(titleLabel);
            Controls.Add(searchLabel);
            Controls.Add(_searchBox);
            Controls.Add(_wordListBox);
            Controls.Add(_countLabel);
            Controls.Add(_deleteButton);
            Controls.Add(btnClearIgnores);
            Controls.Add(btnAddSelection);
            Controls.Add(btnIgnoreSelection);
            Controls.Add(newWordLabel);
            Controls.Add(_newWordBox);
            Controls.Add(btnAdd);
            Controls.Add(btnClose);
            Controls.Add(_statusLabel);

            CancelButton = btnClose;

            RefreshWordList();
        }

        private void RefreshWordList()
        {
            var filter = _searchBox.Text.Trim().Normalize(NormalizationForm.FormC);
            var snapshot = PersonalDictionaryManager.Instance.GetUserWordsResult();
            var allWords = snapshot.Words;
            ShowResult(snapshot.Persistence);
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

            var result = PersonalDictionaryManager.Instance.AddUserWord(word);
            ShowResult(result);
            if (!result.Succeeded || result.Status == PersonalDictionaryStatus.Duplicate) return;
            _newWordBox.Clear();
            RefreshWordList();
            _wordListBox.SelectedItem = word.Normalize(NormalizationForm.FormC);
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
                var confirm = MessageBox.Show(this,
                    "Xóa mục đang chọn khỏi Từ điển cá nhân?",
                    "Từ điển cá nhân", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);
                if (confirm != DialogResult.Yes) return;
                var result = PersonalDictionaryManager.Instance.RemoveUserWord(word);
                ShowResult(result);
                if (!result.Succeeded) return;
                RefreshWordList();
            }
        }

        private void AddSelectedText()
        {
            if (string.IsNullOrWhiteSpace(_selectedText)) return;
            var result = PersonalDictionaryManager.Instance.AddUserWord(_selectedText);
            ShowResult(result);
            if (!result.Succeeded || result.Status == PersonalDictionaryStatus.Duplicate) return;
            _searchBox.Clear();
            RefreshWordList();
            _wordListBox.SelectedItem = _selectedText;
        }

        private void IgnoreSelectedText()
        {
            if (string.IsNullOrWhiteSpace(_selectedText) || string.IsNullOrWhiteSpace(_currentDocumentId)) return;
            ShowResult(PersonalDictionaryManager.Instance.IgnoreWordForDocument(
                _currentDocumentId, _selectedText));
        }

        private void ClearDocumentIgnores()
        {
            if (string.IsNullOrWhiteSpace(_currentDocumentId))
            {
                var clearAll = MessageBox.Show(this,
                    "Không có dữ liệu quét của tài liệu hiện tại. Xóa danh sách bỏ qua tạm thời của toàn bộ phiên Word?",
                    "Danh sách bỏ qua", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (clearAll == DialogResult.Yes)
                    ShowResult(PersonalDictionaryManager.Instance.ClearAllDocumentIgnores());
                return;
            }

            var result = MessageBox.Show(this,
                "Chọn Có để xóa danh sách bỏ qua của tài liệu hiện tại.\n" +
                "Chọn Không để xóa danh sách bỏ qua của toàn bộ phiên Word.\n" +
                "Chọn Hủy để giữ nguyên.",
                "Danh sách bỏ qua", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
                ShowResult(PersonalDictionaryManager.Instance.ClearDocumentIgnores(_currentDocumentId));
            else if (result == DialogResult.No)
                ShowResult(PersonalDictionaryManager.Instance.ClearAllDocumentIgnores());
        }

        private void ShowResult(PersonalDictionaryResult result)
        {
            _statusLabel.Text = result.Status == PersonalDictionaryStatus.Duplicate
                ? result.Message
                : result.Succeeded ? string.Empty : result.Message;
        }

        private static string? NormalizeOptionalEntry(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            return value!.Trim(' ', '\t', '\r', '\n', '\a').Normalize(NormalizationForm.FormC);
        }

        public static void Prompt(IWin32Window? owner = null, string? currentDocumentId = null,
            string? selectedText = null)
        {
            using (var dialog = new CustomDictionaryDialog(currentDocumentId, selectedText))
            {
                ((Form)dialog).ShowDialog(owner);
            }
        }
    }
}
