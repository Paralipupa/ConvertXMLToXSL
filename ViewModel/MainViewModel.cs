using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using Egrn.Model;
using Egrn.Properties;

namespace Egrn.ViewModel
{
    class MainViewModel : ObservableObject
    {
        #region Property

        private CancellationTokenSource _cts = null;

        public StatusBar StatusPanel { get; set; } = new StatusBar();

        private bool _isWait;
        public bool IsWait
        {
            get { return _isWait; }
            set
            {
                _isWait = value;
                OnPropertyChanged("IsWait");
            }
        }

        public string PatternText
        {
            get { return Properties.Settings.Default.patternFile; }
            set
            {
                Properties.Settings.Default.patternFile = value;
                OnPropertyChanged("PatternText");
                Properties.Settings.Default.Save();
            }
        }

        public string XSLText
        {
            get { return Properties.Settings.Default.xslFile; }
            set
            {
                Properties.Settings.Default.xslFile = value;
                OnPropertyChanged("XSLText");
                Properties.Settings.Default.Save();
            }
        }

        public string InputText
        {
            get { return Properties.Settings.Default.inputFile; }
            set
            {
                Properties.Settings.Default.inputFile = value;
                OnPropertyChanged("InputText");
                Properties.Settings.Default.Save();
            }
        }

        public string OutputText
        {
            get { return Properties.Settings.Default.outputFile; }
            set
            {
                Properties.Settings.Default.outputFile = value;
                OnPropertyChanged("OutputText");
                Properties.Settings.Default.Save();
            }
        }

        public string PathZip
        {
            get { return Properties.Settings.Default.pathZip; }
            set
            {
                Properties.Settings.Default.pathZip = value;
                OnPropertyChanged("PathZip");
                Properties.Settings.Default.Save();
            }
        }

        public string PathXml
        {
            get { return Properties.Settings.Default.pathXml; }
            set
            {
                Properties.Settings.Default.pathXml = value;
                OnPropertyChanged("PathXml");
                Properties.Settings.Default.Save();
            }
        }

        public string PathPDF
        {
            get { return Properties.Settings.Default.pathPDF; }
            set
            {
                Properties.Settings.Default.pathPDF = value;
                OnPropertyChanged("PathPDF");
                Properties.Settings.Default.Save();
            }
        }

        public string PathExcel
        {
            get { return Properties.Settings.Default.pathExcel; }
            set
            {
                Properties.Settings.Default.pathExcel = value;
                OnPropertyChanged("PathExcel");
                Properties.Settings.Default.Save();
            }
        }


        public bool IsExtractPDF
        {
            get { return bool.Parse(Properties.Settings.Default.isExtractPDF); }
            set
            {
                Properties.Settings.Default.isExtractPDF = value.ToString();
                OnPropertyChanged("IsExtractPDF");
                Properties.Settings.Default.Save();
            }
        }

        public bool IsExtractExcel
        {
            get { return bool.Parse(Properties.Settings.Default.isExtractExcel); }
            set
            {
                Properties.Settings.Default.isExtractExcel = value.ToString();
                OnPropertyChanged("IsExtractExcel");
                Properties.Settings.Default.Save();
            }
        }

        public bool IsUnionXML
        {
            get { return bool.Parse(Properties.Settings.Default.isUnionXML); }
            set
            {
                Properties.Settings.Default.isUnionXML = value.ToString();
                OnPropertyChanged("IsUnionXML");
                Properties.Settings.Default.Save();
            }
        }

        #endregion

        #region Command


        public ICommand SelectPattCommand => new Command(
            _ =>
            {
                DialogService dlg = new DialogService();
                dlg.File = PatternText;
                if (dlg.OpenFileDialog("xls(*.xls)|*.xls*|All files (*.*)|*.*"))
                {
                    PatternText = dlg.File;
                }
            });

        public ICommand SelectInputCommand => new Command(
            _ =>
            {
                DialogService dlg = new DialogService();
                dlg.File = InputText;
                if (dlg.OpenFileDialog())
                {
                    InputText = dlg.File;
                }
            });

        public ICommand SelectXSLCommand => new Command(
             _ =>
             {
                 DialogService dlg = new DialogService();
                 dlg.File = XSLText;
                 if (dlg.OpenFileDialog("xsl(*.xsl)|*.xsl*|All files (*.*)|*.*"))
                 {
                     XSLText = dlg.File;
                 }
             });

        public ICommand SelectOutputCommand => new Command(
            _ =>
            {
                DialogService dlg = new DialogService();
                if (dlg.OpenFileDialog("xls(*.xls)|*.xls*|All files (*.*)|*.*"))
                {
                    OutputText = dlg.File;
                }
            });

        public ICommand SelectPathZipCommand => new Command(
             _ =>
             {
                 DialogService dlg = new DialogService();
                 dlg.Path = PathZip;
                 if (dlg.OpenPathDialog())
                 {
                     PathZip = dlg.Path;
                 }
             });

        public ICommand SelectPathXmlCommand => new Command(
              _ =>
              {
                  DialogService dlg = new DialogService();
                  dlg.Path = PathXml;
                  if (dlg.OpenPathDialog())
                  {
                      PathXml = dlg.Path;
                  }
              });

        public ICommand SelectPathPDFCommand => new Command(
              _ =>
              {
                  DialogService dlg = new DialogService();
                  dlg.Path = PathPDF;
                  if (dlg.OpenPathDialog())
                  {
                      PathPDF = dlg.Path;
                  }
              });
        
        public ICommand SelectPathExcelCommand => new Command(
              _ =>
              {
                  DialogService dlg = new DialogService();
                  dlg.Path = PathExcel;
                  if (dlg.OpenPathDialog())
                  {
                      PathExcel = dlg.Path;
                  }
              });

        public ICommand CancelCommand => new Command(
            _ =>
            {
                _cts?.Cancel();
            },
            _ => { return IsWait; });

        public ICommand RunParsingCommand => new Command(
             _ =>
             {
                 Task task = RunParsingAsync();
             },
             _ => { return IsWait == false; });

        public ICommand RunUnzipCommand => new Command(
            _ =>
            {
                Task task = RunUnzipAllAsync();
            },
            _ => { return IsWait == false; });

        public ICommand DeleteAllExceptXMLCommand => new Command(
            _ =>
            {
                Task task = DeleteAllAsync();
            },
            _ => { return IsWait == false; });

        #endregion

        private async Task RunUnzipAllAsync()
        {
            if (Directory.Exists(PathZip) == false) return;
            if (Directory.Exists(PathXml) == false)
            {
                Directory.CreateDirectory(PathXml);
            };

            IsWait = true;
            try
            {
                _cts = new CancellationTokenSource();
                CancellationToken token = _cts.Token;
                var progressindicator = new Progress<ProgressIndicate>(ReportProgress);

                await Task.Run(() =>
                {

                    XmlToXLS xtx = new XmlToXLS();

                    xtx.UnZipAll(PathZip, PathXml, progressindicator, token);
                    xtx.UnZipAll(PathXml, PathXml, progressindicator, token);

                });
            }
            finally
            {
                ReportProgress();
                IsWait = false;

            }

        }

        private async Task RunParsingAsync()
        {

            if (File.Exists(PatternText) == false)
            {
                DialogService.MsgBox($"Шаблон не найден {PatternText}");
                return;
            }
            IsWait = true;
            try
            {
                _cts = new CancellationTokenSource();
                CancellationToken token = _cts.Token;
                var progressindicator = new Progress<ProgressIndicate>(ReportProgress);

                await Task.Run(() =>
                {


                    XmlToXLS xtx = new XmlToXLS();
                    xtx.PathPDF = PathPDF;
                    xtx.PathExcel = PathExcel;
                    xtx.IsExtractToPDF = IsExtractPDF;
                    xtx.IsExtractToExcel = IsExtractExcel;
                    xtx.IsUnionXMLs = IsUnionXML;
                    xtx.ParsingXMLToXLS(PatternText, PathXml, XSLText,  progressindicator, token);

                });
            }
            finally
            {
                ReportProgress();
                IsWait = false;

            }

        }

        private async Task DeleteAllAsync()
        {
            if (Directory.Exists(PathXml) == false)
            {
                DialogService.MsgBox($"Каталог не найден {PathXml}");
                return;
            }

            IsWait = true;
            try
            {
                _cts = new CancellationTokenSource();
                CancellationToken token = _cts.Token;
                var progressindicator = new Progress<ProgressIndicate>(ReportProgress);

                await Task.Run(() =>
                {


                    XmlToXLS xtx = new XmlToXLS();

                    xtx.DeleteAll(PathXml, progressindicator, token, "*.xml");


                });
            }
            finally
            {
                ReportProgress();
                IsWait = false;

            }
        }



        private void ReportProgress(ProgressIndicate progressindicate = null)
        {
            if (progressindicate != null)
            {
                if (progressindicate.GetTotal() != 0) StatusPanel.Progress.Value = progressindicate.GetCurrent() * 100 / progressindicate.GetTotal();

                StatusPanel.Status = $"{StatusPanel.Progress.Value}%";
                StatusPanel.Information = (progressindicate?.GetTitle() ?? "") + " " + (progressindicate?.GetName() ?? "") + " " + (progressindicate.GetCurrent() != 0 ? progressindicate.GetCurrent().ToString() : "");

                if (progressindicate.Message != null)
                {
                    //listRows.Items.Add(progressindicate.Message);
                    progressindicate.ClearMessage();
                }


            }
            else
            {
                StatusPanel.Progress.Value = 0;
                StatusPanel.Information = "";
                StatusPanel.Status = "";
            }
        }

        public MainViewModel()
        {

        }
    }
}
