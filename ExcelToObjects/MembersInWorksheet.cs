using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToObjects {
    public class MembersInWorksheet {
        public List<Member> Members { get; set; }
        public string NewWorksheetName { get; set; }

        public MembersInWorksheet() {
            Members = new List<Member>();
            NewWorksheetName = null;
        }
    }
}
