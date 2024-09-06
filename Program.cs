using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: off2pdf.exe <path-to-file>");
            return;
        }

        string filePath = args[0];
        if (!File.Exists(filePath))
        {
            Console.WriteLine("File not found: " + filePath);
            return;
        }

        string ext = Path.GetExtension(filePath).ToLower();

        switch (ext)
        {
            case ".doc":
            case ".docx":
                ConvertWordToPdf(filePath);
                break;
            case ".ppt":
            case ".pptx":
                ConvertPowerPointToPdf(filePath);
                break;
            case ".xls":
            case ".xlsx":
                ConvertExcelToPdf2(filePath);
                break;
            default:
                Console.WriteLine("Unsupported file type: " + ext);
                break;
        }
    }

    private static void ConvertWordToPdf(string filePath)
    {
        Microsoft.Office.Interop.Word.Application wordApp = null;
        Microsoft.Office.Interop.Word.Document document = null;

        try
        {
            wordApp = new Microsoft.Office.Interop.Word.Application();
            document = wordApp.Documents.Open(filePath);

            string pdfPath = Path.ChangeExtension(filePath, ".pdf");

            document.ExportAsFixedFormat(pdfPath, WdExportFormat.wdExportFormatPDF);

            Console.WriteLine($"Converted Word to PDF: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            if (document != null)
            {
                document.Close(false);
                Marshal.ReleaseComObject(document);
            }

            if (wordApp != null)
            {
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static void ConvertPowerPointToPdf(string filePath)
    {
        Microsoft.Office.Interop.PowerPoint.Application pptApp = null;
        Presentation presentation = null;

        try
        {
            pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
            presentation = pptApp.Presentations.Open(filePath, WithWindow: MsoTriState.msoFalse);

            string pdfPath = Path.ChangeExtension(filePath, ".pdf");

            presentation.SaveAs(pdfPath, PpSaveAsFileType.ppSaveAsPDF);

            Console.WriteLine($"Converted PowerPoint to PDF: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            if (presentation != null)
            {
                presentation.Close();
                Marshal.ReleaseComObject(presentation);
            }

            if (pptApp != null)
            {
                pptApp.Quit();
                Marshal.ReleaseComObject(pptApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static void ConvertExcelToPdf(string filePath)
    {
        Microsoft.Office.Interop.Excel.Application excelApp = null;
        Workbook workbook = null;

        try
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            workbook = excelApp.Workbooks.Open(filePath);

            string pdfPath = Path.ChangeExtension(filePath, ".pdf");

            workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfPath);

            Console.WriteLine($"Converted Excel to PDF: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            if (workbook != null)
            {
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }

            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
    private static void ConvertExcelToPdf2(string filePath)
    {
        Microsoft.Office.Interop.Excel.Application excelApp = null;
        Workbook workbook = null;

        try
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            workbook = excelApp.Workbooks.Open(filePath);

            string pdfPath = Path.ChangeExtension(filePath, ".pdf");

            workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfPath, Quality: XlFixedFormatQuality.xlQualityStandard);

            Console.WriteLine($"Converted Excel to PDF: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            if (workbook != null)
            {
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
            }

            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

}

