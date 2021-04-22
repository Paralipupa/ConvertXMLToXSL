using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Egrn.Model
{
    class ProgressIndicate
    {
        private string _name;
        private string _title;
        private readonly int _total;
        private string _message;
        private int _current;
        private int _interval;

        public string Message => _message;

        public ProgressIndicate() : this(0)
        {
        }

        public ProgressIndicate(int total)
        {
            _current = 0;
            _interval = 0;
            _title = null;
            _message = null;
            _name = null;
            _total = total;
        }

        public void AddMessage(string str) { _message = str; }

        public void ClearMessage() { _message = null; }

        public void ClearInterval() => _interval = 0;

        public string GetName() { return _name; }

        public void SetName(string value) { _name = value; }

        public string GetTitle() { return _title; }

        public void SetTitle(string value) { _title = value; }

        public int GetInterval() { return _interval; }

        public int GetCurrent() { return _current; }

        public int GetTotal() { return _total; }

        public void MoveNextCurrent() { ++_current; ++_interval; }
    }

}
