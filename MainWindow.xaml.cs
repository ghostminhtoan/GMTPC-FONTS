using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
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

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                await Task.Run(() => InstallFonts());
                SetProgress(100, "Completed");
                await Task.Delay(700);
                Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                SetProgress(100, ex.Message);
                await Task.Delay(2500);
                Application.Current.Shutdown(1);
            }
        }

        private void InstallFonts()
        {
            string archivePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GMTPC-FONTS.zip");
            if (!File.Exists(archivePath))
            {
                throw new FileNotFoundException("Font package not found.", archivePath);
            }

            string extractRoot = Path.Combine(Path.GetTempPath(), "GMTPC-FONTS-" + Guid.NewGuid().ToString("N"));

            try
            {
                SetProgress(5, "Extracting");
                ZipFile.ExtractToDirectory(archivePath, extractRoot, Encoding.UTF8);

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
                    InstallFont(fontFiles[i], userFontsDirectory);
                    int progress = 10 + (int)Math.Round(((i + 1) * 85.0) / fontFiles.Count);
                    SetProgress(progress, "Installing");
                }

                BroadcastFontChange();
                SetProgress(98, "Cleaning up");
            }
            finally
            {
                TryDeleteDirectory(extractRoot);
                TryDeleteFile(archivePath);
            }
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
            Dispatcher.Invoke(() =>
            {
                InstallProgress.Value = Math.Max(0, Math.Min(100, value));
                ProgressText.Text = text;
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
