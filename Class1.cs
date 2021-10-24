using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace minsquare
{
    public class JsonResult
    {
        public string rows { get; set; }
        public string name { get; set; }
        public string value { get; set; }

        
    }


    public class Col
    {
        public string id { get; set; }
        public string label { get; set; }
        public string type { get; set; }
        public string pattern { get; set; }
    }

    public class C
    {
        public double v { get; set; }
        public string f { get; set; }
    }

    public class Row
    {
        public List<C> c { get; set; }
    }

    public class Table
    {
        public List<Col> cols { get; set; }
        public List<Row> rows { get; set; }
        public int parsedNumHeaders { get; set; }
    }

    public class RootObject
    {
        public string version { get; set; }
        public string reqId { get; set; }
        public string status { get; set; }
        public string sig { get; set; }
        public Table table { get; set; }
        public List<Table> result { get; set; }
    }
}
