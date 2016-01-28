using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtParser
{
    class ObjectInfo
    {
        string _name;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
        string _linenumer;

        public string LineNumber
        {
            get { return _linenumer; }
            set { _linenumer = value; }
        }

    }
}
