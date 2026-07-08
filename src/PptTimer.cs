using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace PptTimer
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length > 0 && args[0] == "--ppt-test")
            {
                SlideStatus status = PowerPointReader.Read();
                Console.WriteLine(status.IsPresenting
                    ? "PRESENTING " + status.Current + "/" + status.Total
                    : "NOT_PRESENTING " + PowerPointReader.LastError);
                Environment.Exit(status.IsPresenting ? 0 : 1);
                return;
            }

            if (args.Length > 0 && args[0] == "--self-test")
            {
                Environment.Exit(SelfTest.Run());
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new OverlayForm());
        }
    }

    internal sealed class Countdown
    {
        public int DurationSeconds { get; private set; }
        public int RemainingSeconds { get; private set; }
        public bool IsRunning { get; private set; }

        public Countdown(int seconds)
        {
            SetDuration(seconds);
        }

        public void SetDuration(int seconds)
        {
            DurationSeconds = Math.Max(1, seconds);
            RemainingSeconds = DurationSeconds;
            IsRunning = false;
        }

        public void StartFromBeginning()
        {
            RemainingSeconds = DurationSeconds;
            IsRunning = true;
        }

        public void Clear()
        {
            RemainingSeconds = 0;
            IsRunning = false;
        }

        public void Tick()
        {
            if (!IsRunning)
                return;
            RemainingSeconds--;
        }
    }

    internal sealed class AppSettings
    {
        public int DurationSeconds = 600;
        public int WarningSeconds = 60;

        private static string FilePath
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PPT倒计时设置.ini"); }
        }

        public static AppSettings Load()
        {
            AppSettings result = new AppSettings();
            try
            {
                foreach (string line in File.ReadAllLines(FilePath, Encoding.UTF8))
                {
                    string[] parts = line.Split(new[] { '=' }, 2);
                    int value;
                    if (parts.Length != 2 || !int.TryParse(parts[1], out value))
                        continue;
                    if (parts[0] == "duration" && value > 0)
                        result.DurationSeconds = value;
                    if (parts[0] == "warning" && value >= 0)
                        result.WarningSeconds = value;
                }
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            return result;
        }

        public void Save()
        {
            File.WriteAllText(
                FilePath,
                "duration=" + DurationSeconds + Environment.NewLine +
                "warning=" + WarningSeconds + Environment.NewLine,
                new UTF8Encoding(false));
        }
    }

    internal struct SlideStatus
    {
        public bool IsPresenting;
        public int Current;
        public int Total;
    }

    internal static class PowerPointReader
    {
        public static string LastError { get; private set; }

        public static SlideStatus Read()
        {
            object app = null;
            object windows = null;
            object window = null;
            object view = null;
            object slide = null;
            object presentation = null;
            object slides = null;

            try
            {
                LastError = "";
                app = Marshal.GetActiveObject("PowerPoint.Application");
                windows = Get(app, "SlideShowWindows");
                int count = Convert.ToInt32(Get(windows, "Count"));
                if (count < 1)
                    return new SlideStatus();

                window = GetIndexed(windows, 1);
                view = Get(window, "View");
                slide = Get(view, "Slide");
                presentation = Get(window, "Presentation");
                slides = Get(presentation, "Slides");

                return new SlideStatus
                {
                    IsPresenting = true,
                    Current = Convert.ToInt32(Get(slide, "SlideIndex")),
                    Total = Convert.ToInt32(Get(slides, "Count"))
                };
            }
            catch (COMException ex) { LastError = ex.Message; return new SlideStatus(); }
            catch (InvalidCastException ex) { LastError = ex.Message; return new SlideStatus(); }
            catch (TargetInvocationException ex)
            {
                LastError = ex.InnerException == null ? ex.Message : ex.InnerException.Message;
                return new SlideStatus();
            }
            finally
            {
                Release(slides);
                Release(presentation);
                Release(slide);
                Release(view);
                Release(window);
                Release(windows);
                Release(app);
            }
        }

        private static object Get(object target, string property)
        {
            return target.GetType().InvokeMember(
                property, BindingFlags.GetProperty, null, target, null);
        }

        private static object GetIndexed(object target, int index)
        {
            return target.GetType().InvokeMember(
                "Item", BindingFlags.GetProperty | BindingFlags.InvokeMethod, null, target, new object[] { index });
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                Marshal.FinalReleaseComObject(value);
        }
    }

    internal sealed class OverlayForm : Form
    {
        private const int WsExNoActivate = 0x08000000;
        private readonly Color background = Color.FromArgb(241, 239, 235);
        private readonly Color card = Color.FromArgb(255, 255, 255);
        private readonly Color primary = Color.FromArgb(24, 24, 29);
        private readonly Color accent = Color.FromArgb(39, 68, 190);
        private readonly Color warning = Color.FromArgb(218, 139, 36);
        private readonly Color danger = Color.FromArgb(216, 68, 82);
        private readonly Color muted = Color.FromArgb(75, 72, 82);

        private readonly Timer secondTimer = new Timer();
        private readonly Timer pptTimer = new Timer();
        private readonly Button settingsButton;
        private readonly Button aboutButton;
        private readonly Button closeButton;
        private readonly AppSettings settings;
        private readonly Countdown countdown;
        private SlideStatus slideStatus;
        private bool wasPresenting;
        private Point dragOrigin;

        public OverlayForm()
        {
            settings = AppSettings.Load();
            countdown = new Countdown(settings.DurationSeconds);

            Text = "PPT 演讲计时器";
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ClientSize = new Size(267, 86);
            BackColor = background;
            TopMost = true;
            Opacity = 0.78;
            ShowInTaskbar = true;
            DoubleBuffered = true;
            Font = new Font("Microsoft YaHei UI", 9F);

            settingsButton = MakeButton("⚙", 212, 3, 18, Color.Transparent, accent);
            settingsButton.Font = new Font("Segoe UI Symbol", 10F);
            aboutButton = MakeButton("?", 231, 3, 18, Color.Transparent, accent);
            aboutButton.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            closeButton = MakeButton("×", 249, 3, 18, Color.Transparent, muted);
            closeButton.Font = new Font("Microsoft YaHei UI", 10F);
            settingsButton.Click += ShowSettings;
            aboutButton.Click += ShowAbout;
            closeButton.Click += delegate { Close(); };
            Controls.Add(settingsButton);
            Controls.Add(aboutButton);
            Controls.Add(closeButton);

            secondTimer.Interval = 1000;
            secondTimer.Tick += delegate { countdown.Tick(); RefreshState(); };
            secondTimer.Start();

            pptTimer.Interval = 500;
            pptTimer.Tick += delegate
            {
                slideStatus = PowerPointReader.Read();
                if (slideStatus.IsPresenting && !wasPresenting)
                    countdown.StartFromBeginning();
                else if (!slideStatus.IsPresenting && wasPresenting)
                    countdown.Clear();
                wasPresenting = slideStatus.IsPresenting;
                RefreshState();
            };
            pptTimer.Start();

            MouseDown += BeginDrag;
            MouseMove += Drag;
            MouseUp += EndDrag;
            Resize += delegate { SetRoundedRegion(); };
            Shown += delegate
            {
                Rectangle area = Screen.PrimaryScreen.WorkingArea;
                Location = new Point(area.Right - Width - 24, area.Top + 24);
                SetRoundedRegion();
            };
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams parameters = base.CreateParams;
                parameters.ExStyle |= WsExNoActivate;
                return parameters;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            DrawRoundedCard(g, new Rectangle(7, 7, 145, 72), 12);
            DrawRoundedCard(g, new Rectangle(158, 7, 102, 72), 12);

            bool overtime = slideStatus.IsPresenting && countdown.RemainingSeconds < 0;
            bool showTime = !overtime || ((Environment.TickCount / 500) % 2 == 0);
            Color timeColor = !slideStatus.IsPresenting
                ? muted
                : (countdown.RemainingSeconds < 0
                    ? danger
                    : (countdown.RemainingSeconds <= settings.WarningSeconds ? warning : primary));
            string time = slideStatus.IsPresenting ? FormatTime(countdown.RemainingSeconds) : "--:--";
            if (showTime)
            {
                using (Font f = new Font("Segoe UI", 26F, FontStyle.Bold))
                using (SolidBrush b = new SolidBrush(timeColor))
                {
                    SizeF size = g.MeasureString(time, f);
                    g.DrawString(time, f, b, 79.5F - size.Width / 2F, 18);
                }
            }

            string slideText = "--/--";
            if (slideStatus.IsPresenting)
            {
                int remaining = Math.Max(0, slideStatus.Total - slideStatus.Current);
                slideText = remaining + "/" + slideStatus.Total;
            }
            using (Font f = new Font("Segoe UI", 18F, FontStyle.Bold))
            using (SolidBrush b = new SolidBrush(slideStatus.IsPresenting ? accent : muted))
            {
                SizeF size = g.MeasureString(slideText, f);
                g.DrawString(slideText, f, b, 209 - size.Width / 2F, 31);
            }
        }

        private Button MakeButton(string text, int x, int y, int width, Color back, Color fore)
        {
            Button button = new Button();
            button.Text = text;
            button.Location = new Point(x, y);
            button.Size = new Size(width, 27);
            button.BackColor = back;
            button.ForeColor = fore;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(229, 232, 244);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(216, 221, 240);
            button.Cursor = Cursors.Hand;
            button.TabStop = false;
            return button;
        }

        private void ShowSettings(object sender, EventArgs e)
        {
            using (SettingsForm form = new SettingsForm(
                countdown.DurationSeconds, settings.WarningSeconds, background, card, primary, accent, muted))
            {
                if (form.ShowDialog(this) != DialogResult.OK)
                    return;
                settings.DurationSeconds = form.DurationSeconds;
                settings.WarningSeconds = form.WarningSeconds;
                settings.Save();
                countdown.SetDuration(settings.DurationSeconds);
                if (slideStatus.IsPresenting)
                    countdown.StartFromBeginning();
                RefreshState();
            }
        }

        private void ShowAbout(object sender, EventArgs e)
        {
            using (AboutForm form = new AboutForm(background, card, primary, accent, muted))
                form.ShowDialog(this);
        }

        private void RefreshState()
        {
            Invalidate();
        }

        private void DrawRoundedCard(Graphics g, Rectangle bounds, int radius)
        {
            using (GraphicsPath path = RoundedPath(bounds, radius))
            using (SolidBrush b = new SolidBrush(card))
                g.FillPath(b, path);
        }

        private void SetRoundedRegion()
        {
            using (GraphicsPath path = RoundedPath(ClientRectangle, 18))
                Region = new Region(path);
        }

        private static GraphicsPath RoundedPath(Rectangle r, int radius)
        {
            int d = radius * 2;
            GraphicsPath path = new GraphicsPath();
            path.AddArc(r.Left, r.Top, d, d, 180, 90);
            path.AddArc(r.Right - d - 1, r.Top, d, d, 270, 90);
            path.AddArc(r.Right - d - 1, r.Bottom - d - 1, d, d, 0, 90);
            path.AddArc(r.Left, r.Bottom - d - 1, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private static string FormatTime(int seconds)
        {
            string sign = seconds < 0 ? "-" : "";
            int value = Math.Abs(seconds);
            return sign + (value / 60).ToString("00") + ":" + (value % 60).ToString("00");
        }

        private void BeginDrag(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
                dragOrigin = e.Location;
        }

        private void Drag(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
                Location = new Point(Left + e.X - dragOrigin.X, Top + e.Y - dragOrigin.Y);
        }

        private void EndDrag(object sender, MouseEventArgs e) { }
    }

    internal sealed class SettingsForm : Form
    {
        private readonly NumericUpDown minutes;
        private readonly NumericUpDown seconds;
        private readonly NumericUpDown warning;
        private readonly Color primaryColor;

        public int DurationSeconds
        {
            get { return (int)minutes.Value * 60 + (int)seconds.Value; }
        }

        public int WarningSeconds
        {
            get { return (int)warning.Value; }
        }

        public SettingsForm(int duration, int warningSeconds, Color background, Color card, Color primary, Color accent, Color muted)
        {
            primaryColor = primary;
            Text = "设置倒计时";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(360, 270);
            BackColor = background;
            ForeColor = primary;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Font = new Font("Microsoft YaHei UI", 9F);

            Label title = LabelAt("设置倒计时", 24, 22, 220, 28, primary, 15F, true);
            Label durationLabel = LabelAt("时长", 24, 72, 80, 24, muted, 9F, false);
            Label minuteUnit = LabelAt("分", 152, 106, 24, 24, muted, 9F, false);
            Label secondUnit = LabelAt("秒", 277, 106, 24, 24, muted, 9F, false);
            Label quickLabel = LabelAt("快捷", 24, 143, 44, 24, muted, 9F, false);
            Label warningLabel = LabelAt("剩余多少秒时变为橙色", 24, 181, 190, 24, muted, 9F, false);
            Label warningUnit = LabelAt("秒", 277, 181, 24, 24, muted, 9F, false);

            minutes = NumberAt(24, 101, 120, 29, 0, 999, duration / 60, card);
            seconds = NumberAt(178, 101, 90, 29, 0, 59, duration % 60, card);
            warning = NumberAt(215, 176, 53, 29, 0, 9999, warningSeconds, card);

            Button quick10 = QuickButton("10分钟", 72, 138, 72, card, primary, 600);
            Button quick20 = QuickButton("20分钟", 154, 138, 72, card, primary, 1200);
            Button quick30 = QuickButton("30分钟", 236, 138, 72, card, primary, 1800);
            Button cancel = DialogButton("取消", 190, 226, 68, card, primary);
            Button save = DialogButton("保存", 268, 226, 68, accent, Color.White);
            cancel.DialogResult = DialogResult.Cancel;
            save.DialogResult = DialogResult.OK;
            save.Click += delegate
            {
                if (DurationSeconds < 1)
                {
                    MessageBox.Show(this, "倒计时时长至少为 1 秒。", "无法保存", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DialogResult = DialogResult.None;
                }
            };

            Controls.AddRange(new Control[]
            {
                title, durationLabel, minuteUnit, secondUnit, quickLabel, warningLabel, warningUnit,
                minutes, seconds, warning, quick10, quick20, quick30, cancel, save
            });
            AcceptButton = save;
            CancelButton = cancel;
        }

        private Label LabelAt(string text, int x, int y, int w, int h, Color color, float size, bool bold)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                ForeColor = color,
                Font = new Font("Microsoft YaHei UI", size, bold ? FontStyle.Bold : FontStyle.Regular)
            };
        }

        private NumericUpDown NumberAt(int x, int y, int w, int h, int min, int max, int value, Color back)
        {
            return new NumericUpDown
            {
                Location = new Point(x, y),
                Size = new Size(w, h),
                Minimum = min,
                Maximum = max,
                Value = Math.Min(max, Math.Max(min, value)),
                BackColor = back,
                ForeColor = primaryColor,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private Button DialogButton(string text, int x, int y, int w, Color back, Color fore)
        {
            return new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, 30),
                BackColor = back,
                ForeColor = fore,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.None
            };
        }

        private Button QuickButton(string text, int x, int y, int w, Color back, Color fore, int totalSeconds)
        {
            Button button = DialogButton(text, x, y, w, back, fore);
            button.Click += delegate
            {
                minutes.Value = totalSeconds / 60;
                seconds.Value = totalSeconds % 60;
            };
            return button;
        }
    }

    internal sealed class AboutForm : Form
    {
        public AboutForm(Color background, Color card, Color primary, Color accent, Color muted)
        {
            Text = "关于";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(390, 360);
            BackColor = background;
            ForeColor = primary;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Font = new Font("Microsoft YaHei UI", 9F);

            Label title = new Label
            {
                Text = "PPT 演讲计时器",
                Location = new Point(24, 20),
                Size = new Size(260, 30),
                ForeColor = accent,
                Font = new Font("Microsoft YaHei UI", 16F, FontStyle.Bold)
            };
            Label meta = new Label
            {
                Text = "版本 1.1\r\n作者 Nick Lee\r\n邮箱 shaou77@sina.com",
                Location = new Point(25, 58),
                Size = new Size(310, 72),
                ForeColor = primary,
                Font = new Font("Microsoft YaHei UI", 10F)
            };
            Label art = new Label
            {
                Text =
                    "      .-''''-.\r\n" +
                    "    .'        '.\r\n" +
                    "   /   .----.   \\\r\n" +
                    "  |   /      \\   |\r\n" +
                    "  |  |  @--@  |  |\r\n" +
                    "  |  |   __   |  |\r\n" +
                    "  |   \\ '--' /   |\r\n" +
                    "   \\    '--'    /\r\n" +
                    "    '.  .__.  .'\r\n" +
                    "      '-.__.-'\r\n" +
                    "       /|  |\\\r\n" +
                    "      /_|__|_\\",
                Location = new Point(25, 130),
                Size = new Size(320, 172),
                BackColor = card,
                ForeColor = muted,
                Font = new Font("Consolas", 8F),
                Padding = new Padding(12)
            };
            Button ok = new Button
            {
                Text = "确定",
                Location = new Point(298, 316),
                Size = new Size(68, 30),
                BackColor = accent,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.OK
            };
            ok.FlatAppearance.BorderSize = 0;

            Controls.AddRange(new Control[] { title, meta, art, ok });
            AcceptButton = ok;
            CancelButton = ok;
        }
    }

    internal static class SelfTest
    {
        public static int Run()
        {
            int failures = 0;
            failures += Check("初始时长", delegate
            {
                Countdown c = new Countdown(61);
                return c.RemainingSeconds == 61 && !c.IsRunning;
            });
            failures += Check("开始并计时", delegate
            {
                Countdown c = new Countdown(2);
                c.StartFromBeginning();
                c.Tick();
                return c.RemainingSeconds == 1 && c.IsRunning;
            });
            failures += Check("到零后继续负计时", delegate
            {
                Countdown c = new Countdown(1);
                c.StartFromBeginning();
                c.Tick();
                c.Tick();
                return c.RemainingSeconds == -1 && c.IsRunning;
            });
            failures += Check("退出放映清空", delegate
            {
                Countdown c = new Countdown(5);
                c.StartFromBeginning();
                c.Tick();
                c.Clear();
                return c.RemainingSeconds == 0 && !c.IsRunning;
            });
            failures += Check("放映开始自动重置并计时", delegate
            {
                Countdown c = new Countdown(5);
                c.StartFromBeginning();
                return c.RemainingSeconds == 5 && c.IsRunning;
            });
            failures += Check("非法时长保护", delegate
            {
                Countdown c = new Countdown(0);
                return c.DurationSeconds == 1;
            });

            Console.WriteLine(failures == 0 ? "全部自检通过。" : failures + " 项自检失败。");
            return failures == 0 ? 0 : 1;
        }

        private static int Check(string name, Func<bool> test)
        {
            try
            {
                bool passed = test();
                Console.WriteLine((passed ? "[通过] " : "[失败] ") + name);
                return passed ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("[异常] " + name + ": " + ex.Message);
                return 1;
            }
        }
    }
}
