namespace SpreadsheetEditor.Controls.Printing;

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using SkiaSharp;

// Windows-only printing via the native common print dialog (comdlg32 PrintDlg) + GDI. Shows the
// standard Windows print dialog so the user picks a printer/settings, then prints by blitting the
// pre-rendered page bitmaps onto the printer's device context. Self-contained: it does not depend on
// any registered PDF handler. All P/Invoke is guarded behind an OS check by the caller.
[SupportedOSPlatform("windows")]
internal static class WindowsPrinter
{
    // Shows the print dialog and prints the given page bitmaps. Returns true if printing was started,
    // false if the user cancelled or the dialog/printer context could not be created.
    public static bool PrintPages(IntPtr ownerHwnd, string documentTitle, IReadOnlyList<SKBitmap> pages)
    {
        if (pages is null || pages.Count == 0)
        {
            return false;
        }

        PRINTDLG pd = default;
        pd.lStructSize = Marshal.SizeOf<PRINTDLG>();
        pd.hwndOwner = ownerHwnd;
        pd.Flags = PD_RETURNDC | PD_NOPAGENUMS | PD_NOSELECTION | PD_USEDEVMODECOPIESANDCOLLATE;
        pd.nCopies = 1;

        if (!PrintDlg(ref pd))
        {
            return false; // user cancelled (or the dialog failed to open)
        }

        try
        {
            if (pd.hDC == IntPtr.Zero)
            {
                return false;
            }

            int copies = Math.Max((int)pd.nCopies, 1);
            int pageWidth = GetDeviceCaps(pd.hDC, HORZRES);
            int pageHeight = GetDeviceCaps(pd.hDC, VERTRES);

            DOCINFO info = new()
            {
                cbSize = Marshal.SizeOf<DOCINFO>(),
                lpszDocName = string.IsNullOrWhiteSpace(documentTitle) ? "Spreadsheet" : documentTitle,
            };

            if (StartDoc(pd.hDC, ref info) <= 0)
            {
                return false;
            }

            SetStretchBltMode(pd.hDC, HALFTONE);
            for (int copy = 0; copy < copies; copy++)
            {
                foreach (SKBitmap page in pages)
                {
                    if (StartPage(pd.hDC) <= 0)
                    {
                        EndDoc(pd.hDC);
                        return true;
                    }

                    BlitPage(pd.hDC, page, pageWidth, pageHeight);
                    EndPage(pd.hDC);
                }
            }

            EndDoc(pd.hDC);
            return true;
        }
        finally
        {
            if (pd.hDC != IntPtr.Zero)
            {
                DeleteDC(pd.hDC);
            }

            if (pd.hDevMode != IntPtr.Zero)
            {
                GlobalFree(pd.hDevMode);
            }

            if (pd.hDevNames != IntPtr.Zero)
            {
                GlobalFree(pd.hDevNames);
            }
        }
    }

    // Stretches one BGRA (top-down) page bitmap to fill the printable area of the device context.
    private static void BlitPage(IntPtr hdc, SKBitmap page, int destWidth, int destHeight)
    {
        if (page.ColorType != SKColorType.Bgra8888)
        {
            using SKBitmap converted = new(new SKImageInfo(page.Width, page.Height, SKColorType.Bgra8888, SKAlphaType.Premul));
            page.CopyTo(converted, SKColorType.Bgra8888);
            Stretch(hdc, converted, destWidth, destHeight);
            return;
        }

        Stretch(hdc, page, destWidth, destHeight);
    }

    private static void Stretch(IntPtr hdc, SKBitmap bitmap, int destWidth, int destHeight)
    {
        BITMAPINFOHEADER header = new()
        {
            biSize = Marshal.SizeOf<BITMAPINFOHEADER>(),
            biWidth = bitmap.Width,
            biHeight = -bitmap.Height, // negative => top-down (matches Skia's pixel order)
            biPlanes = 1,
            biBitCount = 32,
            biCompression = 0, // BI_RGB
        };

        StretchDIBits(
            hdc,
            0, 0, destWidth, destHeight,
            0, 0, bitmap.Width, bitmap.Height,
            bitmap.GetPixels(), ref header, DIB_RGB_COLORS, SRCCOPY);
    }

    // ----- Win32 interop -----

    private const int PD_RETURNDC = 0x00000100;

    private const int PD_USEDEVMODECOPIESANDCOLLATE = 0x00040000;

    private const int PD_NOSELECTION = 0x00000004;

    private const int PD_NOPAGENUMS = 0x00000008;

    private const int HORZRES = 8;

    private const int VERTRES = 10;

    private const int HALFTONE = 4;

    private const int DIB_RGB_COLORS = 0;

    private const uint SRCCOPY = 0x00CC0020;

    [StructLayout(LayoutKind.Sequential)]
    private struct PRINTDLG
    {
        public int lStructSize;
        public IntPtr hwndOwner;
        public IntPtr hDevMode;
        public IntPtr hDevNames;
        public IntPtr hDC;
        public int Flags;
        public short nFromPage;
        public short nToPage;
        public short nMinPage;
        public short nMaxPage;
        public short nCopies;
        public IntPtr hInstance;
        public IntPtr lCustData;
        public IntPtr lpfnPrintHook;
        public IntPtr lpfnSetupHook;
        public IntPtr lpPrintTemplateName;
        public IntPtr lpSetupTemplateName;
        public IntPtr hPrintTemplate;
        public IntPtr hSetupTemplate;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct DOCINFO
    {
        public int cbSize;
        [MarshalAs(UnmanagedType.LPWStr)] public string lpszDocName;
        [MarshalAs(UnmanagedType.LPWStr)] public string? lpszOutput;
        [MarshalAs(UnmanagedType.LPWStr)] public string? lpszDatatype;
        public int fwType;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct BITMAPINFOHEADER
    {
        public int biSize;
        public int biWidth;
        public int biHeight;
        public short biPlanes;
        public short biBitCount;
        public int biCompression;
        public int biSizeImage;
        public int biXPelsPerMeter;
        public int biYPelsPerMeter;
        public int biClrUsed;
        public int biClrImportant;
    }

    [DllImport("comdlg32.dll", EntryPoint = "PrintDlgW", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PrintDlg(ref PRINTDLG lppd);

    [DllImport("gdi32.dll", EntryPoint = "StartDocW", SetLastError = true)]
    private static extern int StartDoc(IntPtr hdc, ref DOCINFO lpdi);

    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern int StartPage(IntPtr hdc);

    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern int EndPage(IntPtr hdc);

    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern int EndDoc(IntPtr hdc);

    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern int GetDeviceCaps(IntPtr hdc, int index);

    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern int SetStretchBltMode(IntPtr hdc, int mode);

    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern int StretchDIBits(
        IntPtr hdc,
        int xDest, int yDest, int destWidth, int destHeight,
        int xSrc, int ySrc, int srcWidth, int srcHeight,
        IntPtr bits, ref BITMAPINFOHEADER bmi, int usage, uint rop);

    [DllImport("gdi32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteDC(IntPtr hdc);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalFree(IntPtr hMem);
}
