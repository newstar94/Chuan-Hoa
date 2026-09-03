using System;
using System.Drawing;
using System.Windows.Forms;

namespace VietDocStandardizer
{
    public class TaskPaneControl : UserControl
    {
        private Button btnAutoFix;
        private ComboBox cboRegime;
        private Label lblRegime;
        private Label lblStatus;
        private Button btnCleanSpaces;
        private Button btnRepeatHeaders;
        private Button btnPreventSplit;
        private Button btnToneMark;
        private Panel pnlHeader;
        private Label lblTitle;

        public TaskPaneControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.pnlHeader = new Panel();
            this.lblTitle = new Label();
            this.lblRegime = new Label();
            this.cboRegime = new ComboBox();
            this.btnAutoFix = new Button();
            this.btnCleanSpaces = new Button();
            this.btnRepeatHeaders = new Button();
            this.btnPreventSplit = new Button();
            this.btnToneMark = new Button();
            this.lblStatus = new Label();

            this.SuspendLayout();

            // Header
            this.pnlHeader.BackColor = Color.FromArgb(37, 99, 235);
            this.pnlHeader.Dock = DockStyle.Top;
            this.pnlHeader.Height = 50;
            this.pnlHeader.Controls.Add(this.lblTitle);

            this.lblTitle.Text = "⚡ CHUẨN HÓA THỂ THỨC";
            this.lblTitle.ForeColor = Color.White;
            this.lblTitle.Font = new Font("Segoe UI", 11, FontStyle.Bold);
            this.lblTitle.Location = new Point(12, 14);
            this.lblTitle.AutoSize = true;

            // Regime Label & Combobox
            this.lblRegime.Text = "CHẾ ĐỘ THỂ THỨC:";
            this.lblRegime.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            this.lblRegime.Location = new Point(14, 65);
            this.lblRegime.AutoSize = true;

            this.cboRegime.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cboRegime.Items.AddRange(new object[] {
                "Nghị định 30/2020/NĐ-CP (Nhà nước)",
                "Hướng dẫn 05-HD/VPTW (Văn bản Đảng)",
                "Quy chế Doanh nghiệp Viettel"
            });
            this.cboRegime.SelectedIndex = 0;
            this.cboRegime.Location = new Point(14, 88);
            this.cboRegime.Width = 320;
            this.cboRegime.SelectedIndexChanged += (s, e) => {
                if (cboRegime.SelectedIndex == 0) WordInteropEngine.CurrentRegime = "ND30";
                else if (cboRegime.SelectedIndex == 1) WordInteropEngine.CurrentRegime = "DANG_HD05";
                else WordInteropEngine.CurrentRegime = "VIETTEL";
            };

            // 1-Click AutoFix Button
            this.btnAutoFix.Text = "✨ 1-CLICK AUTO-FIX TOÀN DIỆN";
            this.btnAutoFix.BackColor = Color.FromArgb(37, 99, 235);
            this.btnAutoFix.ForeColor = Color.White;
            this.btnAutoFix.FlatStyle = FlatStyle.Flat;
            this.btnAutoFix.Font = new Font("Segoe UI", 10, FontStyle.Bold);
            this.btnAutoFix.Location = new Point(14, 125);
            this.btnAutoFix.Size = new Size(320, 45);
            this.btnAutoFix.Click += (s, e) => WordInteropEngine.Perform1ClickAutoFix();

            // Additional Action Buttons
            this.btnCleanSpaces.Text = "🧹 Dọn khoảng trắng & dấu câu";
            this.btnCleanSpaces.Location = new Point(14, 185);
            this.btnCleanSpaces.Size = new Size(320, 32);
            this.btnCleanSpaces.Click += (s, e) => WordInteropEngine.CleanSpacesAndPunctuation();

            this.btnToneMark.Text = "🔤 Chuẩn hóa dấu thanh (hòa, thúy)";
            this.btnToneMark.Location = new Point(14, 225);
            this.btnToneMark.Size = new Size(320, 32);
            this.btnToneMark.Click += (s, e) => WordInteropEngine.NormalizeToneMarks();

            this.btnRepeatHeaders.Text = "📑 Lặp hàng tiêu đề bảng khi tràn trang";
            this.btnRepeatHeaders.Location = new Point(14, 265);
            this.btnRepeatHeaders.Size = new Size(320, 32);
            this.btnRepeatHeaders.Click += (s, e) => WordInteropEngine.RepeatTableHeaders();

            this.btnPreventSplit.Text = "🛡️ Chống rách bảng giữa 2 trang";
            this.btnPreventSplit.Location = new Point(14, 305);
            this.btnPreventSplit.Size = new Size(320, 32);
            this.btnPreventSplit.Click += (s, e) => WordInteropEngine.PreventTableRowSplitting();

            // Status Label
            this.lblStatus.Text = "Trạng thái: Sẵn sàng xử lý";
            this.lblStatus.ForeColor = Color.DarkGray;
            this.lblStatus.Font = new Font("Segoe UI", 8);
            this.lblStatus.Location = new Point(14, 350);
            this.lblStatus.AutoSize = true;

            // Form setup
            this.Controls.Add(this.pnlHeader);
            this.Controls.Add(this.lblRegime);
            this.Controls.Add(this.cboRegime);
            this.Controls.Add(this.btnAutoFix);
            this.Controls.Add(this.btnCleanSpaces);
            this.Controls.Add(this.btnToneMark);
            this.Controls.Add(this.btnRepeatHeaders);
            this.Controls.Add(this.btnPreventSplit);
            this.Controls.Add(this.lblStatus);

            this.Font = new Font("Segoe UI", 9);
            this.Size = new Size(350, 600);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
