using PrintPdfFile;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDFConverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RawPrinterHelper.SendFileToPrinter(@"A:\output\test.pdf");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // generate a file name as the current date/time in unix timestamp format
            string file = (string)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds.ToString();
            // the directory to store the output.
            var directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // initialize PrinterDocument object
            PrintDocument doc = new PrintDocument()
            {
                PrinterSettings = new PrinterSettings()
                {
                    // set the printer to 'Microsoft Print to PDF'
                    PrinterName = "Microsoft Print to PDF",
                    // tell the object this document will print to file
                    PrintToFile = true,
                    // set the filename to whatever you like (full path)
                    PrintFileName = Path.Combine(directory, file + ".pdf"),
                }

            };

        }

        private void button3_Click(object sender, EventArgs e)
        {
            PrintFile(@"A:\output\test.docx");
        }
        /// <summary>
        /// Convert a file to PDF using office _Document object
        /// </summary>
        /// <param name="InputFile">Full path and filename with extension of the file you want to convert from</param>
        /// <returns></returns>
        public void PrintFile(string InputFile)
        {
            // convert input filename to new pdf name
            object OutputFileName = Path.Combine(
                Path.GetDirectoryName(InputFile),
                Path.GetFileNameWithoutExtension(InputFile) + ".pdf"
            );


            // Set an object so there is less typing for values not needed
            object missing = System.Reflection.Missing.Value;
 
            //// `doc` is of type `_Document`
            doc.PrintOut(
                ref missing,    // Background
                ref missing,    // Append
                ref missing,    // Range
                OutputFileName, // OutputFileName
                ref missing,    // From
                ref missing,    // To
                ref missing,    // Item
                ref missing,    // Copies
                ref missing,    // Pages
                ref missing,    // PageType
                ref missing,    // PrintToFile
                ref missing,    // Collate
                ref missing,    // ActivePrinterMacGX
                ref missing,    // ManualDuplexPrint
                ref missing,    // PrintZoomColumn
                ref missing,    // PrintZoomRow
                ref missing,    // PrintZoomPaperWidth
                ref missing,    // PrintZoomPaperHeight
            );
        }
    }
}
