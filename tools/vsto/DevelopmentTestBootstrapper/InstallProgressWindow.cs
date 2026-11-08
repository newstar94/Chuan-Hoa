using System;
using System.Drawing;
using System.Windows.Forms;

namespace ChuanHoa.DevelopmentTestBootstrapper
{
    /// <summary>
    /// Progress window shown between the installer's confirmation and completion
    /// messages. The install transaction runs on this UI thread, so the window is
    /// refreshed at each step boundary. Every size derives from the system message
    /// font, which the per-monitor DPI-aware manifest renders at the real display
    /// DPI: the window stays sharp and proportionate at 100 % to 200 % scaling.
    /// It cannot be closed by the user: an interrupted transaction would have to
    /// roll back anyway.
    /// </summary>
    internal sealed class InstallProgressWindow : IDisposable
    {
        // Finer than the step count so the final visual-style animation is short.
        private const int UnitsPerStep = 10;

        private readonly Form _form;
        private readonly Label _status;
        private readonly ProgressBar _bar;
        private readonly Font _headingFont;
        private readonly int _steps;
        private int _started;
        private bool _closing;
        private bool _disposed;

        internal InstallProgressWindow(string title, string heading, int steps)
        {
            _steps = Math.Max(1, steps);
            var font = SystemFonts.MessageBoxFont ?? Control.DefaultFont;
            var unit = font.Height;
            var width = unit * 28;
            _headingFont = new Font(font, FontStyle.Bold);

            _form = new Form
            {
                Text = title,
                Font = font,
                AutoScaleMode = AutoScaleMode.None,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                ControlBox = false,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = true,
                StartPosition = FormStartPosition.CenterScreen,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                // Form padding applies to docked children only; it adds the right
                // and bottom margins, the layout's location the left and top ones.
                Padding = new Padding(0, 0, unit, unit)
            };
            var layout = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 3,
                Location = new Point(unit, unit),
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };
            var headingLabel = new Label
            {
                Text = heading,
                Font = _headingFont,
                AutoSize = true,
                MaximumSize = new Size(width, 0),
                Margin = new Padding(0, 0, 0, unit / 2)
            };
            _status = new Label
            {
                AutoSize = false,
                Size = new Size(width, unit * 2),
                Margin = new Padding(0, 0, 0, unit / 2)
            };
            _bar = new ProgressBar
            {
                Size = new Size(width, unit * 6 / 5),
                Minimum = 0,
                Maximum = _steps * UnitsPerStep,
                Style = ProgressBarStyle.Continuous,
                Margin = Padding.Empty
            };
            layout.Controls.Add(headingLabel, 0, 0);
            layout.Controls.Add(_status, 0, 1);
            layout.Controls.Add(_bar, 0, 2);
            _form.Controls.Add(layout);
            _form.FormClosing += (sender, args) =>
            {
                if (!_closing && args.CloseReason == CloseReason.UserClosing) args.Cancel = true;
            };

            _form.Show();
            _form.Activate();
            Pump();
        }

        /// <summary>Shows that the next step started; the bar counts the steps already finished.</summary>
        internal void Begin(string status)
        {
            if (_disposed) return;
            SetValue(Math.Min(_started, _steps) * UnitsPerStep);
            _started++;
            _status.Text = status;
            Pump();
        }

        /// <summary>Replaces the status text without advancing (for example while rolling back).</summary>
        internal void Status(string status)
        {
            if (_disposed) return;
            _status.Text = status;
            Pump();
        }

        internal void Complete(string status)
        {
            if (_disposed) return;
            SetValue(_bar.Maximum);
            _status.Text = status;
            Pump();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _closing = true;
            _form.Close();
            _form.Dispose();
            _headingFont.Dispose();
        }

        private void SetValue(int value)
        {
            // Visual styles animate forward slowly but move backward at once:
            // overshoot by one unit and step back so the bar shows the real value.
            value = Math.Max(_bar.Minimum, Math.Min(_bar.Maximum, value));
            if (value < _bar.Maximum)
            {
                _bar.Value = value + 1;
                _bar.Value = value;
            }
            else
            {
                _bar.Value = value;
                _bar.Value = value - 1;
                _bar.Value = value;
            }
        }

        private void Pump()
        {
            _form.Refresh();
            Application.DoEvents();
        }
    }
}
