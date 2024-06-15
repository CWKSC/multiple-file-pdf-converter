# multiple_file_pdf_converter

> [!NOTE]
> 
> Created when I am in Higher Diploma
> 
> Teacher gives a lot of doc and ppt, open those file is slow
> 
> pdf can open in web browser, this is much convenient
> 

> [!NOTE]
> 
> The project is old (2020), I just recreated to update project from:
> 
> https://github.com/CWKSC/multithread_pdf_converter_backup
> 
> In fact I don't think it really run in multithread, so I rename it to multiple_file in this repo

## Expected output

```
[1 / 3] aaa.pdf
[2 / 3] bbb.pdf
[3 / 3] ccc.pdf

All work finsih! Spent 7.5423706 seconds
Press any Enter to exit ...
```

## Source code

```csharp
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace multiple_file_pdf_converter
{
    public class FileDialog
    {
        public static string[] GetFilepaths()
        {
            var fileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = "Please select doc, docx, ppt, pptx files that need to be converted to pdf",
                Filter = "(*.doc, *.docx, *.ppt, *pptx)|*.doc;*.docx;*.ppt;*.pptx"
            };

            var isOk = fileDialog.ShowDialog();
            fileDialog.Dispose();
            if (isOk != DialogResult.OK) { Environment.Exit(1); }

            return fileDialog.FileNames;
        }
    }

    public class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var filepaths = FileDialog.GetFilepaths();
            totalWork = filepaths.Length;

            var stopwatch = Stopwatch.StartNew();

            Task[] tasks = new Task[totalWork];

            for (int i = 0; i < totalWork; i++)
            {
                string filepath = filepaths[i];

                string extension = System.IO.Path.GetExtension(filepath);

                string sourcePath = filepath;
                string targetPath = filepath.Substring(0, filepath.Length - extension.Length) + ".pdf";

                if (IsWord(extension))
                {
                    tasks[i] = Task.Run(() => WordToPDF(sourcePath, targetPath));
                }
                else if (IsPowerPoint(extension))
                {
                    tasks[i] = Task.Run(() => PowerPointToPDF(sourcePath, targetPath));
                }
            }

            Task.WhenAll(tasks).Wait();

            stopwatch.Stop();
            Console.WriteLine("\nAll work finsih! Spent " + stopwatch.Elapsed.TotalSeconds + " seconds");
            Console.Write("Press any Enter to exit ...");
            Console.ReadLine();
        }

        public static bool IsWord(string extension) => extension.Equals(".doc") || extension.Equals(".docx");
        public static bool IsPowerPoint(string extension) => extension.Equals(".ppt") || extension.Equals(".pptx");

        public static void WordToPDF(string sourcePath, string targetPath)
        {
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Document document = null;
            try
            {
                application.Visible = false;
                document = application.Documents.Open(sourcePath);
                document.ExportAsFixedFormat(targetPath, WdExportFormat.wdExportFormatPDF);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                FinishedWorkAddOne_ShowProgress(targetPath);
                document.Close();
                application.Quit();
            }
        }

        public static void PowerPointToPDF(string sourcePath, string targetPath)
        {
            Microsoft.Office.Interop.PowerPoint.Application application = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation presentation = application.Presentations.Open(sourcePath, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
            try
            {
                presentation.ExportAsFixedFormat(targetPath, PpFixedFormatType.ppFixedFormatTypePDF);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                FinishedWorkAddOne_ShowProgress(targetPath);
                presentation.Close();
                application.Quit();
            }
        }


        public static int totalWork = 0;

        public static int finishedWorkNumber = 0;
        public static readonly object Lock = new object();
        public static void FinishedWorkAddOne_ShowProgress(string targetPath)
        {
            lock (Lock)
            {
                finishedWorkNumber++;
                Console.WriteLine($"[{finishedWorkNumber} / {totalWork}] {targetPath}");
            }
        }

    }
}
```
