using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtParser
{
    public class DataObjectBinding
    {
        public DataObjectBinding(string fileName, string className, string methodName, string methodLineNumber, string storeProcedureName, string storedProcedureLineNumber)
        {
            FileName = fileName;
            ClassName = className;
            MethodName = methodName;
            MethodLineNumber = methodLineNumber;
            StoredProcedureName = storeProcedureName;
            StoredProcedureLineNumber = storedProcedureLineNumber;
        }

        public string FileName { get; set; }
        public string ClassName { get; set; }
        public string MethodName { get; set; }
        public string MethodLineNumber { get; set; }
        public string StoredProcedureName { get; set; }
        public string StoredProcedureLineNumber { get; set; }
    }
}
