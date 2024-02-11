using Hardcodet.Wpf.TaskbarNotification;
using Microsoft.Win32;
using NAudio.CoreAudioApi;
using MicMon.Properties;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using NAudio.Wave;
using System.Threading;

namespace MicMon
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private const int IconSize = 32;
        private const int GraphDecayMs = 1500;

        private TaskbarIcon _taskbarIcon;
        private bool _startMinimised = false;

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public string IconPath { get; set; } = default!;

        public double Sensitivity
        {
            get
            {
                return 1.0;
            }
            set
            {
                OnPropertyChanged("LabelContentForGain");
                SaveSettings();
            }
        }

        public string LabelContentForSensitivity => $"Gain: {Sensitivity:N2}";

        private double[] logGraphValues = [0, 0, 0, 0, 0];
        private int undecayedMsAccumulator = 0;
        private void UpdateIcon(int percent, int deltatms)
        {
            byte[] rgbBuffer = new byte[IconSize * IconSize * 4];

            double meter = 0;
            if (percent < 1) { meter = 0; }
            else if (percent < 10) { meter = 0.1; }
            else if (percent < 20) { meter = 0.2; }
            else if (percent < 40) { meter = 0.4; }
            else if (percent < 60) { meter = 0.6; }
            else if (percent < 80) { meter = 0.8; }
            else if (percent < 90) { meter = 0.9; }
            else { meter = 1.0; }

            undecayedMsAccumulator += deltatms;
            if (undecayedMsAccumulator >= GraphDecayMs)
            {
                for (int i = logGraphValues.Length - 1; i > 0; i--)
                {
                    logGraphValues[i] = logGraphValues[i - 1];
                }
                logGraphValues[0] = 0;
                undecayedMsAccumulator = 0;
            }

            if (meter > logGraphValues[0])
                logGraphValues[0] = meter;

            for (int graphTimeSection = 0; graphTimeSection < logGraphValues.Length; graphTimeSection++)
            {
                var sectionWidthPx = IconSize >> (graphTimeSection + 1);
                var graphValue = logGraphValues[graphTimeSection];
                for (int xPx = (sectionWidthPx - 1) * 4; xPx < (2 * sectionWidthPx * 4); xPx += 4)
                {
                    for (int yPx = 0; yPx < IconSize; yPx++)
                    {
                        var rowOffset = yPx * IconSize * 4;
                        rgbBuffer[rowOffset + xPx] = 0;                                         // Blue
                        rgbBuffer[rowOffset + xPx + 1] = graphValue == 0 ? (byte)0 : (byte)128; // Green
                        rgbBuffer[rowOffset + xPx + 2] = (byte)(255 * graphValue);              // Red
                        rgbBuffer[rowOffset + xPx + 3] = 255;                                   // Alpha
                    }
                }
            }

            _taskbarIcon.Icon = IconFromRgbBuffer(rgbBuffer);
        }

        static Icon IconFromRgbBuffer(byte[] rgbBuffer)
        {
            // Create a Bitmap from RGB buffer
            Bitmap bitmap = new Bitmap(IconSize, IconSize, PixelFormat.Format32bppArgb);
            BitmapData bitmapData = bitmap.LockBits(new Rectangle(0, 0, IconSize, IconSize), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
            Marshal.Copy(rgbBuffer, 0, bitmapData.Scan0, rgbBuffer.Length);
            bitmap.UnlockBits(bitmapData);

            var icon = IconHelper.PngIconFromImage(bitmap, IconSize);

            return icon;
        }

        private DateTime _lastIconUpdateTime = DateTime.UtcNow;
        private void WaveIn_DataAvailable(object? sender, WaveInEventArgs e)
        {
            if (DateTime.UtcNow.Subtract(_lastIconUpdateTime).TotalMilliseconds < (GraphDecayMs / 2))
                return;

            var now = DateTime.UtcNow;
            var deltatms = (int)now.Subtract(_lastIconUpdateTime).TotalMilliseconds;
            _lastIconUpdateTime = now;

            // copy buffer into an array of integers
            Int16[] values = new Int16[e.Buffer.Length / 2];
            Buffer.BlockCopy(e.Buffer, 0, values, 0, e.Buffer.Length);

            // Determine the highest value as a fraction of the maximum possible value
            // 16 bit signed int (short) takes values between -32768 and 32767.
            // Avoid negating -32768 (or getting the absolute value) as it's impossible to do inside a 16 bit signed integer.
            float fraction = (float)values.Select(_ => _ == Int16.MinValue ? Int16.MaxValue : Math.Abs(_)).Max() / 32768;
            int percent = (int)(100 * fraction);

            UpdateIcon(percent, deltatms);
        }

        public MainWindow()
        {
            SystemEvents.SessionSwitch += new SessionSwitchEventHandler(SystemEvents_SessionSwitch);

            InitializeComponent();
            DataContext = this;

            _taskbarIcon = (TaskbarIcon)FindResource("NotifyIcon");
            RestartMicMonitoring();

            chkStartAtLogin.IsChecked = (bool)Settings.Default["StartAtLogin"];
            _startMinimised = (bool)Settings.Default["StartMinimised"];
            if (_startMinimised)
            {
                Application.Current.MainWindow.Hide();
            }

            AddCheckboxesForActiveDevices();
        }

        private static WaveInEvent _waveIn = null!;
        private void RestartMicMonitoring()
        {
            try
            {
                if (_waveIn == null)
                {
                    StartMicMonitoring();
                }
                else
                {
                    // If _waveIn is already running then stop it and let the stopped event handler restart it to avoid weird race conditions
                    _waveIn.StopRecording();
                    return;
                }
            }
            catch { }
        }

        private void WaveIn_RecordingStopped(object? sender, StoppedEventArgs e)
        {
            StartMicMonitoring();
        }

        private void StartMicMonitoring()
        {
            var selectedDevicesDocument = ReadSelectedDevicesFromUserSettings();

            var activeDevices = new MMDeviceEnumerator()
                .EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)
                .Select((_, index) => new SelectedDevice(index, _.ID, _.FriendlyName));

            if (_waveIn != null)
            {
                _waveIn.DataAvailable -= WaveIn_DataAvailable;
                _waveIn.RecordingStopped -= WaveIn_RecordingStopped;
                _waveIn.Dispose();
            }

            foreach (var activeDevice in activeDevices)
            {
                var foundActiveSelectedDevice = selectedDevicesDocument.SelectedDevices.FirstOrDefault(_ => _.Id == activeDevice.Id);

                if (foundActiveSelectedDevice != null)
                {
                    _waveIn = new WaveInEvent
                    {
                        DeviceNumber = activeDevice.DeviceNumber,
                        WaveFormat = new WaveFormat(rate: 8000, bits: 16, channels: 1),
                        BufferMilliseconds = 100
                    };

                    _waveIn.DataAvailable += WaveIn_DataAvailable;
                    _waveIn.RecordingStopped += WaveIn_RecordingStopped;
                    _waveIn.StartRecording();

                    break;
                }
            }
        }

        private void AddCheckboxesForActiveDevices()
        {
            var selectedDevicesDocument = ReadSelectedDevicesFromUserSettings();

            var activeDevices = new MMDeviceEnumerator()
                .EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)
                .Select(_ => new SelectedDevice(-1, _.ID, _.FriendlyName));

            panelActiveDevices.Children.Clear();
            foreach (var activeDevice in activeDevices)
            {
                var radioButton = new RadioButton();
                radioButton.Content = activeDevice.FriendlyName;
                radioButton.Margin = new Thickness(3);
                radioButton.IsChecked = selectedDevicesDocument.SelectedDevices.Any(_ => _.Id == activeDevice.Id);
                radioButton.Checked += DeviceCheckbox_CheckedChanged;
                radioButton.Unchecked += DeviceCheckbox_CheckedChanged;
                radioButton.Tag = activeDevice;
                panelActiveDevices.Children.Add(radioButton);
            }
        }

        private void DeviceCheckbox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton radioButton && radioButton.IsChecked.GetValueOrDefault())
            {
                SaveSelectedDevices();
                RestartMicMonitoring();
            }
        }

        private void SaveSettings()
        {
            Settings.Default["Gain"] = 1.0;
            Settings.Default["StartAtLogin"] = chkStartAtLogin.IsChecked;
            Settings.Default["StartMinimised"] = _startMinimised;
            Settings.Default.Save();
        }

        private void SaveSelectedDevices()
        {
            var activeDeviceRadioButtons = panelActiveDevices.Children.OfType<UIElement>().Where(_ => _ is RadioButton).Select(_ => _ as RadioButton);

            var previousSelectedDevicesDocument = ReadSelectedDevicesFromUserSettings();
            var previouslySelectedDevicesThatAreNoLongerActive = previousSelectedDevicesDocument.SelectedDevices.Where(sd => !activeDeviceRadioButtons.Any(adc => (adc!.Tag as SelectedDevice)!.Id == sd.Id)).ToList();

            var newSelectedDevicesDocument = new SelectedDevicesDocument();
            foreach (var radioButton in activeDeviceRadioButtons)
            {
                if (radioButton!.IsChecked.GetValueOrDefault())
                {
                    var selectedDevice = radioButton.Tag as SelectedDevice;
                    newSelectedDevicesDocument.SelectedDevices.Add(new SelectedDevice(-1, selectedDevice!.Id, selectedDevice.FriendlyName));
                }
            }

            // Preserve devices that are currently not shown in the UI, but were selected in the UI before.
            newSelectedDevicesDocument.SelectedDevices = newSelectedDevicesDocument.SelectedDevices.Concat(previouslySelectedDevicesThatAreNoLongerActive).ToList();

            Settings.Default["SelectedDevices"] = JsonSerializer.Serialize(newSelectedDevicesDocument);
            Settings.Default.Save();
        }

        private SelectedDevicesDocument ReadSelectedDevicesFromUserSettings()
        {
            try
            {
                var selectedDevices = (string)Settings.Default["SelectedDevices"];
                selectedDevices = string.IsNullOrEmpty(selectedDevices) ? "{}" : selectedDevices;
                return JsonSerializer.Deserialize<SelectedDevicesDocument>(selectedDevices)!;
            }
            catch
            {
                return new SelectedDevicesDocument();
            }
        }

        private void btnRefreshDevices_Click(object sender, RoutedEventArgs e)
        {
            AddCheckboxesForActiveDevices();
        }

        private void chkStartAtLogin_Checked(object sender, RoutedEventArgs e)
        {
            var path = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
            RegistryKey key = Registry.CurrentUser.OpenSubKey(path, true)!;
            key.SetValue("MicMon", System.Windows.Forms.Application.ExecutablePath.ToString());

            SaveSettings();
        }

        private void chkStartAtLogin_Unchecked(object sender, RoutedEventArgs e)
        {
            var path = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
            RegistryKey key = Registry.CurrentUser.OpenSubKey(path, true)!;
            key.DeleteValue("MicMon", false);

            SaveSettings();
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Minimized)
            {
                Application.Current.MainWindow.Hide();
            }

            _startMinimised = WindowState == WindowState.Minimized;
            SaveSettings();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            WindowState = WindowState.Minimized;
        }

        private void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            if (new[] { SessionSwitchReason.SessionLogon, SessionSwitchReason.SessionUnlock }.Contains(e.Reason))
            {
                RestartMicMonitoring();
            }
        }
    }
}
