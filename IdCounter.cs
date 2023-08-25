using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bs.XML.SpreadSheet {
    public class IdCounter {
        private uint id = 0;
        private readonly string prefix;

        public IdCounter(string prefix = "Id") {
            this.prefix = prefix;
        }
        public override string ToString() {
            return $"{prefix}{id}";
        }
        public static IdCounter operator ++(IdCounter idCounter) {
            IdCounter idc = new IdCounter(idCounter.prefix);
            checked {
                idc.id = ++idCounter.id;
            }
            return idc;
        }
        public static explicit operator uint(IdCounter idCounter) {
            return idCounter.id;
        }
        public static implicit operator string(IdCounter idCounter) {
            return idCounter.ToString();
        }
    }
}
