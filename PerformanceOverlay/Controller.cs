using CommonHelpers;
using ExternalHelpers;
using RTSSSharedMemoryNET;
using System.ComponentModel;

namespace PerformanceOverlay
{
    internal class Controller : IDisposable
    {
        public const String Title = "Performance Overlay";
        public static readonly String TitleWithVersion = Title + " v" + System.Windows.Forms.Application.ProductVersion.ToString();

        Container components = new Container();
        RTSSSharedMemoryNET.OSD? _osd;
        System.Windows.Forms.ContextMenuStrip _contextMenu;
        ToolStripMenuItem _showItem;
        System.Windows.Forms.NotifyIcon _notifyIcon;
        System.Windows.Forms.Timer _osdTimer;
        Sensors _sensors = new Sensors();
        StartupManager _startupManager = new StartupManager(
            Title,
            "Starts Performance Overlay on Windows startup."
        );

        SharedData<OverlayModeSetting> _sharedData = SharedData<OverlayModeSetting>.CreateNew();

        static Controller()
        {
            Dependencies.ValidateRTSSSharedMemoryNET(TitleWithVersion);
        }

        public Controller()
        {
            Instance.OnUninstall(() =>
            {
                _startupManager.Startup = false;
            });

            _contextMenu = new System.Windows.Forms.ContextMenuStrip(components);

            SharedData_Update();
            Instance.Open(TitleWithVersion, Settings.Default.EnableKernelDrivers, "Global\\PerformanceOverlay");
            Instance.RunUpdater(TitleWithVersion);

            if (Instance.WantsRunOnStartup)
                _startupManager.Startup = true;

            var notRunningRTSSItem = _contextMenu.Items.Add("&RTSS is not running");
            notRunningRTSSItem.Enabled = false;
            _contextMenu.Opening += delegate { notRunningRTSSItem.Visible = Dependencies.EnsureRTSS(null) && !OSDHelpers.IsLoaded; };

            _showItem = new ToolStripMenuItem("&Show OSD");
            _showItem.Click += ShowItem_Click;
            _showItem.Checked = Settings.Default.ShowOSD;
            _contextMenu.Items.Add(_showItem);
            _contextMenu.Items.Add(new ToolStripSeparator());
            foreach (var mode in Enum.GetValues<OverlayMode>())
            {
                var modeItem = new ToolStripMenuItem(mode.ToString());
                modeItem.Tag = mode;
                modeItem.Click += delegate
                {
                    Settings.Default.OSDMode = mode;
                    updateContextItems(_contextMenu);
                };
                _contextMenu.Items.Add(modeItem);
            }
            updateContextItems(_contextMenu);

            _contextMenu.Items.Add(new ToolStripSeparator());

            var kernelDriversItem = new ToolStripMenuItem("Use &Kernel Drivers");
            kernelDriversItem.Click += delegate { setKernelDrivers(!Instance.UseKernelDrivers); };
            _contextMenu.Opening += delegate { kernelDriversItem.Checked = Instance.UseKernelDrivers; };
            _contextMenu.Items.Add(kernelDriversItem);

            _contextMenu.Items.Add(new ToolStripSeparator());

            if (_startupManager.IsAvailable)
            {
                var startupItem = new ToolStripMenuItem("Run On Startup");
                startupItem.Checked = _startupManager.Startup;
                startupItem.Click += delegate
                {
                    _startupManager.Startup = !_startupManager.Startup;
                    startupItem.Checked = _startupManager.Startup;
                };
                _contextMenu.Items.Add(startupItem);
            }

            var missingRTSSItem = _contextMenu.Items.Add("&Install missing RTSS");
            missingRTSSItem.Click += delegate { Dependencies.OpenLink(Dependencies.RTSSURL); };
            _contextMenu.Opening += delegate { missingRTSSItem.Visible = !Dependencies.EnsureRTSS(null); };

            var checkForUpdatesItem = _contextMenu.Items.Add("&Check for Updates");
            checkForUpdatesItem.Click += delegate { Instance.RunUpdater(TitleWithVersion, true); };

            var helpItem = _contextMenu.Items.Add("&Help");
            helpItem.Click += delegate { Dependencies.OpenLink(Dependencies.SDTURL); };

            _contextMenu.Items.Add(new ToolStripSeparator());

            var exitItem = _contextMenu.Items.Add("&Exit");
            exitItem.Click += ExitItem_Click;

            _notifyIcon = new System.Windows.Forms.NotifyIcon(components);
            _notifyIcon.Icon = WindowsDarkMode.IsDarkModeEnabled ? Resources.poll_light : Resources.poll;
            _notifyIcon.Text = TitleWithVersion;
            _notifyIcon.Visible = true;
            _notifyIcon.ContextMenuStrip = _contextMenu;

            _osdTimer = new System.Windows.Forms.Timer(components);
            _osdTimer.Tick += OsdTimer_Tick;
            _osdTimer.Interval = 250;
            _osdTimer.Enabled = true;

            if (Settings.Default.ShowOSDShortcut != "")
            {
                GlobalHotKey.RegisterHotKey(Settings.Default.ShowOSDShortcut, () =>
                {
                    Settings.Default.ShowOSD = !Settings.Default.ShowOSD;

                    updateContextItems(_contextMenu);
                });
            }

            if (Settings.Default.CycleOSDShortcut != "")
            {
                GlobalHotKey.RegisterHotKey(Settings.Default.CycleOSDShortcut, () =>
                {
                    var values = Enum.GetValues<OverlayMode>().ToList();

                    int index = values.IndexOf(Settings.Default.OSDMode);
                    Settings.Default.OSDMode = values[(index + 1) % values.Count];
                    Settings.Default.ShowOSD = true;

                    updateContextItems(_contextMenu);
                });
            }

            Microsoft.Win32.SystemEvents.PowerModeChanged += SystemEvents_PowerModeChanged;
        }

        private void SystemEvents_PowerModeChanged(object sender, Microsoft.Win32.PowerModeChangedEventArgs e)
        {
            if (e.Mode == Microsoft.Win32.PowerModes.Resume)
            {
                Instance.HardwareComputer.Reset();
            }
        }

        private void updateContextItems(ContextMenuStrip contextMenu)
        {
            foreach (ToolStripItem item in contextMenu.Items)
            {
                if (item.Tag is OverlayMode)
                    ((ToolStripMenuItem)item).Checked = ((OverlayMode)item.Tag == Settings.Default.OSDMode);
            }

            _showItem.Checked = Settings.Default.ShowOSD;
        }

        private void NotifyIcon_Click(object? sender, EventArgs e)
        {
            if (sender is not null)
            {
                ((NotifyIcon)sender).ContextMenuStrip?.Show(Control.MousePosition);
            }
        }

        private void ShowItem_Click(object? sender, EventArgs e)
        {
            Settings.Default.ShowOSD = !Settings.Default.ShowOSD;
            updateContextItems(_contextMenu);
        }

        private bool AckAntiCheat()
        {
            return AntiCheatSettings.Default.AckAntiCheat(
                TitleWithVersion,
                "Usage of OSD Kernel Drivers might trigger anti-cheat protection in some games.",
                "Ensure that you set it to DISABLED when playing games with ANTI-CHEAT PROTECTION."
            );
        }

        private void setKernelDrivers(bool value)
        {
            if (value && AckAntiCheat())
            {
                Instance.UseKernelDrivers = true;
                Settings.Default.EnableKernelDrivers = true;
            }
            else
            {
                Instance.UseKernelDrivers = false;
                Settings.Default.EnableKernelDrivers = false;
            }
        }

        private void SharedData_Update()
        {
            if (_sharedData.GetValue(out var value))
            {
                if (Enum.IsDefined<OverlayMode>(value.Desired))
                {
                    Settings.Default.OSDMode = (OverlayMode)value.Desired;
                    Settings.Default.ShowOSD = true;
                    updateContextItems(_contextMenu);
                }

                if (Enum.IsDefined<OverlayEnabled>(value.DesiredEnabled))
                {
                    Settings.Default.ShowOSD = (OverlayEnabled)value.DesiredEnabled == OverlayEnabled.Yes;
                    updateContextItems(_contextMenu);
                }

                if (Enum.IsDefined<KernelDriversLoaded>(value.DesiredKernelDriversLoaded))
                {
                    setKernelDrivers((KernelDriversLoaded)value.DesiredKernelDriversLoaded == KernelDriversLoaded.Yes);
                    updateContextItems(_contextMenu);
                }
            }

            _sharedData.SetValue(new OverlayModeSetting()
            {
                Current = Settings.Default.OSDMode,
                CurrentEnabled = Settings.Default.ShowOSD ? OverlayEnabled.Yes : OverlayEnabled.No,
                KernelDriversLoaded = Instance.UseKernelDrivers ? KernelDriversLoaded.Yes : KernelDriversLoaded.No
            });
        }

        private void OsdTimer_Tick(object? sender, EventArgs e)
        {
            try
            {
                _osdTimer.Enabled = false;
                SharedData_Update();
            }
            finally
            {
                _osdTimer.Enabled = true;
            }

            try
            {
                _notifyIcon.Text = TitleWithVersion + ". RTSS Version: " + OSD.Version;
                _notifyIcon.Icon = WindowsDarkMode.IsDarkModeEnabled ? Resources.poll_light : Resources.poll;
            }
            catch
            {
                _notifyIcon.Text = TitleWithVersion + ". RTSS Not Available.";
                _notifyIcon.Icon = Resources.poll_red;
                osdReset();
                return;
            }

            if (!Settings.Default.ShowOSD)
            {
                _osdTimer.Interval = 1000;
                osdReset();
                return;
            }

            _osdTimer.Interval = 250;

            _sensors.Update();

            var osdMode = Settings.Default.OSDMode;

            // If Power Control is visible use temporarily full OSD
            if (Settings.Default.EnableFullOnPowerControl)
            {
                if (SharedData<PowerControlSetting>.GetExistingValue(out var value) && value.Current == PowerControlVisible.Yes)
                    osdMode = OverlayMode.Full;
            }

            var osdOverlay = Overlays.GetOSD(osdMode, _sensors);

            try
            {
                // recreate OSD if not index 0
                if (OSDHelpers.OSDIndex("PerformanceOverlay") != 0)
                    osdClose();
                if (_osd == null)
                    _osd = new OSD("PerformanceOverlay");

                uint offset = 0;
                osdEmbedGraph(ref offset, ref osdOverlay, "[OBJ_FT_SMALL]", -8, -1, 1, 0, 50000.0f, EMBEDDED_OBJECT_GRAPH.FLAG_FRAMETIME);
                osdEmbedGraph(ref offset, ref osdOverlay, "[OBJ_FT_LARGE]", -32, -2, 1, 0, 50000.0f, EMBEDDED_OBJECT_GRAPH.FLAG_FRAMETIME);

                _osd.Update(osdOverlay);
            }
            catch (SystemException)
            {
            }
        }

        private void osdReset()
        {
            try
            {
                if (_osd != null)
                    _osd.Update("");
            }
            catch (SystemException)
            {
            }
        }

        private void osdClose()
        {
            try
            {
                if (_osd != null)
                    _osd.Dispose();
                _osd = null;
            }
            catch (SystemException)
            {
            }
        }

        private void osdEmbedGraph(ref uint offset, ref String osdOverlay, String name, int dwWidth, int dwHeight, int dwMargin, float fltMin, float fltMax, EMBEDDED_OBJECT_GRAPH dwFlags)
        {
            uint size = _osd.EmbedGraph(offset, new float[0], 0, dwWidth, dwHeight, dwMargin, fltMin, fltMax, dwFlags);
            if (size > 0)
                osdOverlay = osdOverlay.Replace(name, "<OBJ=" + offset.ToString("X") + ">");
            offset += size;
        }

        private void ExitItem_Click(object? sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        public void Dispose()
        {
            components.Dispose();
            osdClose();
            using (_sensors) { }
        }
    }
}
