using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CUE.NET;
using CUE.NET.Brushes;
using CUE.NET.Devices.Generic.Enums;
using CUE.NET.Devices.Keyboard;
using CUE.NET.Exceptions;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing.Imaging;

namespace SynthesiaKeyboardLights.Desktop
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        static extern int GetWindowTextLength(IntPtr hWnd);

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern IntPtr GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        const int SRCCOPY = 0xCC0020;

        [DllImport("gdi32.dll")]
        static extern bool BitBlt(
        IntPtr hdcDest, int xDest, int yDest, int width, int height,
        IntPtr hdcSrc, int xSrc, int ySrc, int rop);
        
        CorsairKeyboard keyboard;
        IntPtr SynthesiaWindow;


        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                CueSDK.Initialize();
                Debug.WriteLine("Initialized with " + CueSDK.LoadedArchitecture + "-SDK");

                CorsairKeyboard keyboard = CueSDK.KeyboardSDK ?? throw new WrapperException("No keyboard found");
            }
            catch (CUEException ex)
            {
                Debug.WriteLine("CUE Exception! ErrorCode: " + Enum.GetName(typeof(CorsairError), ex.Error));
            }
            catch (WrapperException ex)
            {
                Debug.WriteLine("Wrapper Exception! Message:" + ex.Message);
            }
            keyboard = CueSDK.KeyboardSDK;
            keyboard.Brush = (SolidColorBrush)Color.Transparent; // This 'enables' manual color changes.
            keyboard.Update();
            SynthesiaWindow = GetWindow();
        }

        private IntPtr GetWindow()
        {
            List<IntPtr> windows = new List<IntPtr>();

            // Enumerate all top-level windows
            EnumWindows((hWnd, lParam) =>
            {
                // Get the length of the window's title
                int length = GetWindowTextLength(hWnd);

                if (length > 0)
                {
                    // Create a buffer to hold the window's title
                    var sb = new StringBuilder(length + 1);

                    // Get the window's title
                    GetWindowText(hWnd, sb, sb.Capacity);

                    // Check if the window's title contains "Synthesia"
                    if (sb.ToString().Contains(" - Synthesia") && sb.ToString() != this.Text && !sb.ToString().Contains(this.Name))
                    {
                        // Add the window to the list
                        windows.Add(hWnd);
                    }
                }

                // Continue enumerating windows
                return true;
            }, IntPtr.Zero);

            // Select the first window containing "Synthesia"
            if (windows.Count > 0)
            {
                return windows[0];
            }
            else
            {
                return IntPtr.Zero;
            }
        }

        public enum PianoKey : uint {
            LeftHandBlackKey = 0xff94b4d6,
            LeftHandWhiteKey = 0xff205082,
            RightHandBlackKey = 0xffabec68,
            RightHandWhiteKey = 0xff4c9800,
        }

        public static unsafe List<Point> FindColorChunks(Bitmap bitmap, PianoKey pk)
        {
            const int targetWidth = 25;
            const int targetHeight = 55;
            int targetColor = unchecked((int)pk);

            BitmapData bitmapData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            List<Point> result = new List<Point>();

            try
            {
                HashSet<int> uniqueXValues = new HashSet<int>(); // to keep track of unique X values

                for (int y = 0; y <= bitmapData.Height - targetHeight; y++)
                {
                    for (int x = 0; x <= bitmapData.Width - targetWidth; x += 25)
                    {
                        int count = 0;
                        for (int dy = 0; dy < targetHeight; dy++)
                        {
                            int* row = (int*)((byte*)bitmapData.Scan0 + (y + dy) * bitmapData.Stride);
                            for (int dx = 0; dx < targetWidth; dx++)
                            {
                                if (row[x + dx] == targetColor)
                                {
                                    count++;
                                }
                                else
                                {
                                    count = 0;
                                    break;
                                }
                            }
                            if (count == targetWidth * targetHeight)
                            {
                                if (uniqueXValues.Add(x)) // only add the x coordinate if it's unique
                                {
                                    result.Add(new Point(x, y));
                                }
                                break;
                            }
                        }
                    }
                }
                return result;
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }
        }

        private void ClearKeys()
        {
            keyboard['A'].Color = Color.Black;
            keyboard['S'].Color = Color.Black;
            keyboard['D'].Color = Color.Black;
            keyboard['F'].Color = Color.Black;
            keyboard['G'].Color = Color.Black;
            keyboard['H'].Color = Color.Black;
            keyboard['J'].Color = Color.Black;
            keyboard['K'].Color = Color.Black;
            keyboard['L'].Color = Color.Black;
            keyboard[CorsairLedId.SemicolonAndColon].Color = Color.Black;
            keyboard[CorsairLedId.ApostropheAndDoubleQuote].Color = Color.Black;
            keyboard['W'].Color = Color.Black;
            keyboard['E'].Color = Color.Black;
            keyboard['T'].Color = Color.Black;
            keyboard['Y'].Color = Color.Black;
            keyboard['U'].Color = Color.Black;
            keyboard['O'].Color = Color.Black;
            keyboard['P'].Color = Color.Black;
        }

        private void TmrPiano_Tick(object sender, EventArgs e)
        {
            if (SynthesiaWindow == null)
            {
                SynthesiaWindow = GetWindow();
            }
            if (SynthesiaWindow == null)
            {
                chkEnabled.Checked = false;
                return;
            }

            // Get the dimensions of the window
            GetWindowRect(SynthesiaWindow, out RECT rect);
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            // Calculate the new height to capture
            int captureHeight = (int)(height * 0.28f);

            if (width == 0 || captureHeight == 0)
            {
                chkEnabled.Checked = false;
                return;
            }

            // Create a new bitmap with the adjusted dimensions
            Bitmap bmp = new Bitmap(width, captureHeight);

            // Get the device context of the window
            IntPtr hdcSrc = GetWindowDC(SynthesiaWindow);

            // Create a graphics object from the bitmap
            using (Graphics g = Graphics.FromImage(bmp))
            {
                // Copy the contents of the window to the bitmap, starting from the bottom
                int captureTop = height - captureHeight;
                IntPtr hdcDest = g.GetHdc();
                BitBlt(hdcDest, 0, 0, width, captureHeight, hdcSrc, 0, captureTop, SRCCOPY);
                g.ReleaseHdc(hdcDest);
            }

            // Release the device context
            ReleaseDC(SynthesiaWindow, hdcSrc);

            ClearKeys();

            foreach (PianoKey pk in (PianoKey[])Enum.GetValues(typeof(PianoKey)))
            {
                List<Point> pianoKeys = FindColorChunks(bmp, pk);

                Color c = ColorTranslator.FromHtml("#4c9800");
                if (pk == PianoKey.RightHandBlackKey)
                {
                    c = ColorTranslator.FromHtml("#abec68");
                }
                else if (pk == PianoKey.LeftHandWhiteKey)
                {
                    c = ColorTranslator.FromHtml("#205082");
                }
                else if (pk == PianoKey.LeftHandBlackKey)
                {
                    c = ColorTranslator.FromHtml("#94b4d6");
                }
                if (pianoKeys.Count > 0)
                {
                    foreach (Point key in pianoKeys)
                    {
                        if (key.X == 25)
                        {
                            keyboard['A'].Color = c;
                        }
                        if (key.X == 125)
                        {
                            keyboard['S'].Color = c;
                        }
                        if (key.X == 225)
                        {
                            keyboard['D'].Color = c;
                        }
                        if (key.X == 300)
                        {
                            keyboard['F'].Color = c;
                        }
                        if (key.X == 400)
                        {
                            keyboard['G'].Color = c;
                        }
                        if (key.X == 475)
                        {
                            keyboard['H'].Color = c;
                        }
                        if (key.X == 575)
                        {
                            keyboard['J'].Color = c;
                        }
                        if (key.X == 650)
                        {
                            keyboard['K'].Color = c;
                        }
                        if (key.X == 750)
                        {
                            keyboard['L'].Color = c;
                        }
                        if (key.X == 850)
                        {
                            keyboard[CorsairLedId.SemicolonAndColon].Color = c;
                        }
                        if (key.X == 925)
                        {
                            keyboard[CorsairLedId.ApostropheAndDoubleQuote].Color = c;
                        }

                        if (key.X == 75)
                        {
                            keyboard['W'].Color = c;
                        }
                        if (key.X == 175)
                        {
                            keyboard['E'].Color = c;
                        }
                        if (key.X == 350)
                        {
                            keyboard['T'].Color = c;
                        }
                        if (key.X == 450)
                        {
                            keyboard['Y'].Color = c;
                        }
                        if (key.X == 550)
                        {
                            keyboard['U'].Color = c;
                        }
                        if (key.X == 700)
                        {
                            keyboard['O'].Color = c;
                        }
                        if (key.X == 800)
                        {
                            keyboard['P'].Color = c;
                        }
                    }
                }
            }
            keyboard.Update();
        }

        private void ChkEnabled_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEnabled.Checked)
            {
                SynthesiaWindow = GetWindow();
            } else
            {
                ClearKeys();
                keyboard.Update();
            }
            tmrPiano.Enabled = chkEnabled.Checked;
        }
    }
}
