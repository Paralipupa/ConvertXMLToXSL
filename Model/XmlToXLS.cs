using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;


namespace Egrn.Model
{
    class XmlToXLS : IDisposable
    {
        private Excel.Application _app = null;
        private bool _firstFile = true;

        public bool IsExtractToPDF { get; set; } = true;
        public bool IsExtractToExcel { get; set; } = false;
        public bool IsUnionXMLs { get; set; } = false;

        public string PathPDF { get; set; }
        public string PathExcel { get; set; }
        public string FileOutput { get; set; }



        /// <summary>
        /// Импортирование данных из файла XML в MS Excel с использованием встроенного в MS Excel механизма XmlMap
        /// </summary>
        /// <param name="xlsFile"></param>
        /// <param name="xmlFile"></param>
        /// <param name="xslFile"></param>
        /// <param name="pathPDF"></param>
        private void Import(string xlsFile, string xmlFile, string xslFile, object wbUnion)
        {
            if (File.Exists(xlsFile) == false)
            {
                return;
            }

            try
            {
                string textFromFile = Xslt.GetXMLString(xmlFile, xslFile);

                Excel.Workbook wb = _app.Workbooks.Open(xlsFile);
                wb.XmlMaps[1].ImportXml(textFromFile);
                Excel.Range range = wb.ActiveSheet.Cells;
                range.Copy();

                Excel.Workbook wbOutput = _app.Workbooks.Add();
                Excel.Range rangeOutput = wbOutput.ActiveSheet.Cells;
                rangeOutput.PasteSpecial();
                int lastRow = rangeOutput.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;


                #region PrintSetup
                _app.PrintCommunication = false;
                wbOutput.ActiveSheet.PageSetup.LeftMargin = _app.InchesToPoints(0);
                wbOutput.ActiveSheet.PageSetup.RightMargin = _app.InchesToPoints(0);
                wbOutput.ActiveSheet.PageSetup.TopMargin = _app.InchesToPoints(0.3);
                wbOutput.ActiveSheet.PageSetup.BottomMargin = _app.InchesToPoints(0.3);
                wbOutput.ActiveSheet.PageSetup.HeaderMargin = _app.InchesToPoints(0.1);
                wbOutput.ActiveSheet.PageSetup.FooterMargin = _app.InchesToPoints(0.1);
                wbOutput.ActiveSheet.PageSetup.FitToPagesWide = 1;
                wbOutput.ActiveSheet.PageSetup.FitToPagesTall = 3;
                _app.PrintCommunication = true;
                #endregion

                string fileName = Path.GetFileName(xmlFile);

                string outputFile = Path.Combine(PathPDF, fileName).Replace(".xml", ".pdf");
                if (IsExtractToPDF)
                {
                    if (File.Exists(outputFile)) File.Delete(outputFile);
                    wbOutput.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputFile);
                }

                if (IsExtractToExcel)
                {
                    outputFile = Path.Combine(PathExcel, fileName).Replace(".xml", ".xlsx");
                    if (File.Exists(outputFile)) File.Delete(outputFile);
                    wbOutput.SaveAs(outputFile);
                }

                if (IsUnionXMLs)
                {
                    lastRow = rangeOutput.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    int lastCol = rangeOutput.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                    Excel.Range rangeIn = rangeOutput.Range[$"A{1 + (_firstFile ? 0 : 1)}", $"{ConvertToLetter(lastCol)}{lastRow }"];


                    Excel.Range rangeUnion = ((Excel.Workbook)wbUnion).ActiveSheet.Cells;
                    lastRow = rangeUnion.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    Excel.Range rangeRowLast = ((Excel.Workbook)wbUnion).Sheets[1].Cells[lastRow + 1, 1];

                    rangeIn.Copy();
                    rangeRowLast.PasteSpecial();

                    _firstFile = false;

                    System.Windows.Application.Current.Dispatcher.BeginInvoke((Action)delegate ()
                    {
                        Clipboard.Clear();
                    });

                }

                wbOutput.Saved = true;
                wbOutput.Close();

                wb.Saved = true;
                wb.Close();

            }
            catch (Exception ex)
            {
                Log.Instance.Write(ex);
            }
        }

        public void ParsingXMLToXLS(string xlsFile, string pathXML, string xslFile, IProgress<ProgressIndicate> progress, CancellationToken token)
        {
            Excel.Workbook wb = null;
            try
            {
                if (IsUnionXMLs)
                {
                    wb = _app.Workbooks.Add();
                }
                string[] dirs = Directory.GetFiles(pathXML, "*.xml");
                ProgressIndicate progressindicate = new ProgressIndicate(dirs.Length);
                foreach (string filename in dirs)
                {
                    if (token.IsCancellationRequested)
                    {
                        throw new Exception("");
                    }

                    progressindicate.SetName(filename);
                    progressindicate.MoveNextCurrent();
                    if (progress != null) progress.Report(progressindicate);

                    Import(xlsFile, filename, xslFile, wb);

                }

            }
            catch (Exception ex)
            {
                if (string.IsNullOrWhiteSpace(ex.Message))
                {
                    Log.Instance.Write(ex);
                }
            }
            finally
            {
                if (wb != null)
                {
                    if (string.IsNullOrWhiteSpace(FileOutput))
                    {
                        FileOutput = Path.Combine(PathExcel, "union.xlsx");
                    }

                    if (File.Exists(FileOutput)) File.Delete(FileOutput);
                    wb.SaveAs(FileOutput);
                    wb.Close();
                }
            }
        }


        public void UnZip(string zipName, string unzipPath)
        {
            try
            {
                ZipFile.ExtractToDirectory(zipName, unzipPath);
            }
            catch (Exception ex)
            {
                Log.Instance.Write(ex);
            }
        }

        public void UnZipAll(string zipPath, string unzipPath, IProgress<ProgressIndicate> progress, CancellationToken token)
        {
            string[] dirs = Directory.GetFiles(zipPath, "*.zip");
            ProgressIndicate progressindicate = new ProgressIndicate(dirs.Length);
            foreach (string filename in dirs)
            {
                if (token.IsCancellationRequested)
                {
                    return;
                }

                progressindicate.SetName(filename);
                progressindicate.MoveNextCurrent();
                if (progress != null) progress.Report(progressindicate);

                UnZip(filename, unzipPath);
            }
        }

        /// <summary>
        /// Удалить все файлы кроме указанных
        /// </summary>
        /// <param name="path"></param>
        /// <param name="exceptionFiles"></param>
        public void DeleteAll(string path, IProgress<ProgressIndicate> progress, CancellationToken token, string filesToIgnore = "")
        {
            IEnumerable<string> dirs = Directory.EnumerateFiles(path);
            int count = 0;
            foreach (var filePath in dirs) count++;

            ProgressIndicate progressindicate = new ProgressIndicate(count);

            foreach (var fileName in dirs)
            {
                if (token.IsCancellationRequested)
                {
                    return;
                }

                progressindicate.SetName(fileName);
                progressindicate.MoveNextCurrent();
                if (progress != null) progress.Report(progressindicate);

                string ext = fileName.Substring(fileName.LastIndexOf('.'));

                if (filesToIgnore.Contains(ext) == false)
                    File.Delete(fileName);
            }
        }


        private string ConvertToLetter(int iCol)
        {

            string nameCol = "";
            while (iCol > 0)
            {
                int a = (iCol - 1) / 26;
                int b = (iCol - 1) % 26;
                nameCol = (Char)(b + 65) + nameCol;
                iCol = a;
            }
            return nameCol;
        }

        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool disposing)
        {
            if (disposing == false) return;
            if (_app != null)
            {
                _app.Quit();
                _app = null;
            }

        }

        ~XmlToXLS()
        {
            Dispose(true);
        }

        #endregion


        public XmlToXLS()
        {
            _app = new Excel.Application();
        }
    }
}
