//https://www.amyuni.com/WebHelp/Amyuni_Document_Converter/Using_the_Developer_Amyuni_Converter.htm

 public class CDIntfCustom
    {
        [DllImport("cdintf.dll", EntryPoint = "DriverInit", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr DriverInit([MarshalAs(UnmanagedType.BStr)] string PrinterName);


        [DllImport("winspool.Drv", EntryPoint = "SetDefaultFileName", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool SetDefaultFileName(IntPtr hPrinter, [MarshalAs(UnmanagedType.LPStr)] string FileName);


        public IntPtr InitializeCustomDriver(string PrinterName)
        {
           return DriverInit(PrinterName);

        }

        public void SetCustomDefaultFileName(IntPtr hPrinter,string FileName)
        {
            SetDefaultFileName(hPrinter, FileName);
        }
    }
