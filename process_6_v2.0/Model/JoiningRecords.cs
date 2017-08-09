using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using process_5_v2._0.Model;
namespace process_6_v2._0.Model
{
    class JoiningRecords
    {
        public Record WC1record { get; set; }
        public claim2Records WC2record { get; set; }
        public JoiningRecords(Record r, claim2Records q)
        {
            this.WC1record = r;
            this.WC2record = q;
        }
    }
}
