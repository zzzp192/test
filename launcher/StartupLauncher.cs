using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace YucaitangReportLauncher
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new StartupForm());
        }
    }

    internal sealed class StartupForm : Form
    {
        private readonly Label _status;
        private readonly Label _elapsed;
        private readonly Timer _timer;
        private readonly Stopwatch _stopwatch;
        private Process _coreProcess;
        private string _coreProcessName;

        public StartupForm()
        {
            Text = "育材堂报告助手正在启动";
            ClientSize = new Size(520, 300);
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.FromArgb(244, 247, 250);
            TopMost = true;
            ShowInTaskbar = true;
            MaximizeBox = false;
            MinimizeBox = false;

            var card = new Panel
            {
                BackColor = Color.White,
                Location = new Point(24, 26),
                Size = new Size(472, 250),
                Padding = new Padding(20)
            };
            Controls.Add(card);

            var accent = new Panel
            {
                BackColor = Color.FromArgb(21, 150, 143),
                Location = new Point(20, 20),
                Size = new Size(72, 6)
            };
            card.Controls.Add(accent);

            var title = new Label
            {
                AutoSize = true,
                Text = "育材堂报告助手",
                Font = new Font("Microsoft YaHei UI", 24, FontStyle.Bold),
                ForeColor = Color.FromArgb(16, 79, 82),
                Location = new Point(20, 52)
            };
            card.Controls.Add(title);

            var subtitle = new Label
            {
                AutoSize = true,
                Text = "材料试验报告处理与 Origin 绘图工具  V3.16",
                Font = new Font("Microsoft YaHei UI", 10),
                ForeColor = Color.FromArgb(71, 85, 105),
                Location = new Point(20, 98)
            };
            card.Controls.Add(subtitle);

            _status = new Label
            {
                AutoSize = true,
                Text = "正在加载核心组件，请稍候…",
                Font = new Font("Microsoft YaHei UI", 10),
                ForeColor = Color.FromArgb(21, 150, 143),
                Location = new Point(20, 148)
            };
            card.Controls.Add(_status);

            var progress = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 24,
                Location = new Point(20, 182),
                Size = new Size(432, 14)
            };
            card.Controls.Add(progress);

            _elapsed = new Label
            {
                AutoSize = true,
                Text = "正在准备启动环境",
                Font = new Font("Microsoft YaHei UI", 8),
                ForeColor = Color.FromArgb(100, 116, 139),
                Location = new Point(20, 212)
            };
            card.Controls.Add(_elapsed);

            _stopwatch = Stopwatch.StartNew();
            _timer = new Timer { Interval = 150 };
            _timer.Tick += CheckCoreWindow;
            Shown += StartCore;
        }

        private void StartCore(object sender, EventArgs e)
        {
            var corePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "育材堂报告助手V3.16_core.exe");
            if (!File.Exists(corePath))
            {
                ShowError("未找到核心程序。请确认启动器与核心程序位于同一文件夹。");
                return;
            }

            try
            {
                _coreProcess = Process.Start(new ProcessStartInfo
                {
                    FileName = corePath,
                    WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory,
                    UseShellExecute = false,
                });
                _coreProcessName = Path.GetFileNameWithoutExtension(corePath);
                _timer.Start();
            }
            catch (Exception exception)
            {
                ShowError("核心程序无法启动：" + exception.Message);
            }
        }

        private void CheckCoreWindow(object sender, EventArgs e)
        {
            var seconds = Math.Max(1, (int)_stopwatch.Elapsed.TotalSeconds);
            _elapsed.Text = string.Format("已加载 {0} 秒，正在准备主界面…", seconds);

            if (_coreProcess == null)
            {
                return;
            }
            if (_coreProcess.HasExited)
            {
                ShowError("核心程序在启动过程中退出。请联系开发者并提供此提示。");
                return;
            }

            // PyInstaller one-file applications create a child process after
            // extraction. Process.Start() returns the extractor parent, so
            // locate the child that owns the finished main window instead of
            // waiting on the parent process for a window it never owns.
            foreach (var candidate in Process.GetProcessesByName(_coreProcessName))
            {
                try
                {
                    candidate.Refresh();
                    if (candidate.MainWindowHandle != IntPtr.Zero &&
                        candidate.MainWindowTitle.Contains("育材堂报告助手 V3.16"))
                    {
                        _timer.Stop();
                        Close();
                        return;
                    }
                }
                finally
                {
                    candidate.Dispose();
                }
            }
        }

        private void ShowError(string message)
        {
            _timer.Stop();
            _status.ForeColor = Color.FromArgb(185, 28, 28);
            _status.Text = "启动失败";
            _elapsed.Text = message;
        }
    }
}
