using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Egrn.Model
{
    class StatusBar : ObservableObject
    {
        public StatusBar()
        {
            _progress = new ProgressBar
            {
                Value = 0,
                Minimum = 0,
                Maximum = 100
            };
        }

        private string _status;
        public string Status
        {
            get { return _status; }
            set { _status = value; OnPropertyChanged("Status"); }
        }
        private string _information;
        public string Information
        {
            get { return _information; }
            set { _information = value; OnPropertyChanged("Information"); }
        }

        private ProgressBar _progress;
        public ProgressBar Progress
        {
            get { return _progress; }
            set { _progress = value; base.OnPropertyChanged("Panel2"); }
        }
    }
}
