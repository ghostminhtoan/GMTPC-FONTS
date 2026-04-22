using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;

namespace GMTPC_FONTS
{
    public partial class MainWindow : Window
    {
        private const int WmFontChange = 0x001D;
        private static readonly IntPtr HwndBroadcast = new IntPtr(0xffff);
        private static readonly string[] FontExtensions = { ".ttf", ".otf", ".ttc", ".fon", ".pfm", ".pfb" };
        private readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        private readonly ManualResetEventSlim pauseGate = new ManualResetEventSlim(true);
        private readonly object statusLock = new object();
        private bool isPaused;
        private string currentStatus = "Preparing";

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                await Task.Run(() => InstallFonts(cancellation.Token));
                SetProgress(100, "Completed");
                SetControlsEnabled(false);
                await Task.Delay(700);
                Application.Current.Shutdown();
            }
            catch (OperationCanceledException)
            {
                SetProgress((int)InstallProgress.Value, "Stopped");
                SetControlsEnabled(false);
                await Task.Delay(900);
                Application.Current.Shutdown(2);
            }
            catch (Exception ex)
            {
                SetProgress(100, ex.Message);
                SetControlsEnabled(false);
                await Task.Delay(2500);
                Application.Current.Shutdown(1);
            }
        }

        private void InstallFonts(CancellationToken token)
        {
            string archivePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GMTPC-FONTS.zip");
            if (!File.Exists(archivePath))
            {
                throw new FileNotFoundException("Font package not found.", archivePath);
            }

            string extractRoot = Path.Combine(Path.GetTempPath(), "GMTPC-FONTS-" + Guid.NewGuid().ToString("N"));

            try
            {
                CheckPauseOrStop(token);
                ExtractArchive(archivePath, extractRoot, token);

                CheckPauseOrStop(token);
                SetProgress(10, "Scanning fonts");
                List<string> fontFiles = FindFontFiles(extractRoot);
                if (fontFiles.Count == 0)
                {
                    throw new InvalidOperationException("No fonts found.");
                }

                string userFontsDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Microsoft",
                    "Windows",
                    "Fonts");
                Directory.CreateDirectory(userFontsDirectory);

                for (int i = 0; i < fontFiles.Count; i++)
                {
                    CheckPauseOrStop(token);
                    SetProgress(10 + (int)Math.Round((i * 85.0) / fontFiles.Count), "Installing " + Path.GetFileName(fontFiles[i]));
                    InstallFont(fontFiles[i], userFontsDirectory);
                    int progress = 10 + (int)Math.Round(((i + 1) * 85.0) / fontFiles.Count);
                    SetProgress(progress, "Installed " + Path.GetFileName(fontFiles[i]));
                }

                CheckPauseOrStop(token);
                SetProgress(97, "Refreshing Windows font list");
                BroadcastFontChange();
                SetProgress(98, "Cleaning up");
            }
            finally
            {
                TryDeleteDirectory(extractRoot);
                if (!token.IsCancellationRequested)
                {
                    TryDeleteFile(archivePath);
                }
            }
        }

        private void ExtractArchive(string archivePath, string extractRoot, CancellationToken token)
        {
            Directory.CreateDirectory(extractRoot);

            using (ZipArchive archive = ZipFile.OpenRead(archivePath))
            {
                int total = archive.Entries.Count;
                for (int i = 0; i < total; i++)
                {
                    CheckPauseOrStop(token);

                    ZipArchiveEntry entry = archive.Entries[i];
                    string destinationPath = GetSafeExtractPath(extractRoot, entry.FullName);
                    int progress = 2 + (int)Math.Round(((i + 1) * 8.0) / Math.Max(1, total));

                    if (string.IsNullOrEmpty(entry.Name))
                    {
                        Directory.CreateDirectory(destinationPath);
                        SetProgress(progress, "Extracting " + entry.FullName);
                        continue;
                    }

                    Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
                    entry.ExtractToFile(destinationPath, true);
                    SetProgress(progress, "Extracting " + entry.Name);
                }
            }
        }

        private static string GetSafeExtractPath(string root, string entryName)
        {
            string destinationPath = Path.GetFullPath(Path.Combine(root, entryName));
            string rootPath = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!destinationPath.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Invalid archive entry.");
            }

            return destinationPath;
        }

        private static List<string> FindFontFiles(string root)
        {
            List<string> result = new List<string>();
            foreach (string file in Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories))
            {
                string extension = Path.GetExtension(file);
                for (int i = 0; i < FontExtensions.Length; i++)
                {
                    if (string.Equals(extension, FontExtensions[i], StringComparison.OrdinalIgnoreCase))
                    {
                        result.Add(file);
                        break;
                    }
                }
            }

            result.Sort(StringComparer.OrdinalIgnoreCase);
            return result;
        }

        private static void InstallFont(string sourcePath, string userFontsDirectory)
        {
            string targetPath = Path.Combine(userFontsDirectory, MakeSafeFileName(Path.GetFileName(sourcePath)));
            if (!File.Exists(targetPath))
            {
                File.Copy(sourcePath, targetPath, false);
            }

            string valueName = Path.GetFileNameWithoutExtension(targetPath) + GetRegistryFontSuffix(targetPath);
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Fonts"))
            {
                if (key != null)
                {
                    key.SetValue(valueName, targetPath, RegistryValueKind.String);
                }
            }

            AddFontResource(targetPath);
        }

        private static string GetRegistryFontSuffix(string path)
        {
            string extension = Path.GetExtension(path);
            if (string.Equals(extension, ".otf", StringComparison.OrdinalIgnoreCase))
            {
                return " (OpenType)";
            }

            return " (TrueType)";
        }

        private static string MakeSafeFileName(string fileName)
        {
            char[] invalidChars = Path.GetInvalidFileNameChars();
            StringBuilder builder = new StringBuilder(fileName.Length);
            foreach (char c in fileName)
            {
                builder.Append(Array.IndexOf(invalidChars, c) >= 0 ? '_' : c);
            }

            return builder.ToString();
        }

        private void SetProgress(int value, string text)
        {
            lock (statusLock)
            {
                currentStatus = text;
            }

            Dispatcher.Invoke(() =>
            {
                InstallProgress.Value = Math.Max(0, Math.Min(100, value));
                ProgressText.Text = isPaused ? "Paused - " + text : text;
            });
        }

        private void CheckPauseOrStop(CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            pauseGate.Wait(token);
            token.ThrowIfCancellationRequested();
        }

        private void PauseButton_Click(object sender, RoutedEventArgs e)
        {
            isPaused = !isPaused;
            if (isPaused)
            {
                pauseGate.Reset();
                PauseButton.Content = "Resume";
            }
            else
            {
                pauseGate.Set();
                PauseButton.Content = "Pause";
            }

            string status;
            lock (statusLock)
            {
                status = currentStatus;
            }

            ProgressText.Text = isPaused ? "Paused - " + status : status;
        }

        private void StopButton_Click(object sender, RoutedEventArgs e)
        {
            StopButton.IsEnabled = false;
            PauseButton.IsEnabled = false;
            pauseGate.Set();
            cancellation.Cancel();
            SetProgress((int)InstallProgress.Value, "Stopping");
        }

        private void SetControlsEnabled(bool enabled)
        {
            Dispatcher.Invoke(() =>
            {
                PauseButton.IsEnabled = enabled;
                StopButton.IsEnabled = enabled;
            });
        }

        private static void TryDeleteDirectory(string path)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                }
            }
            catch
            {
            }
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        private static void BroadcastFontChange()
        {
            UIntPtr result;
            SendMessageTimeout(HwndBroadcast, WmFontChange, UIntPtr.Zero, IntPtr.Zero, 0, 1000, out result);
        }

        [DllImport("gdi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern int AddFontResource(string lpszFilename);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(
            IntPtr hWnd,
            int msg,
            UIntPtr wParam,
            IntPtr lParam,
            int flags,
            int timeout,
            out UIntPtr result);
    }
}
