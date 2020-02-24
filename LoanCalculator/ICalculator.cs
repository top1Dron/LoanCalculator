using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LoanCalculator
{
    public interface ICalculator
    {
        String Calculate();

        void WordExport(object oMissing, object oEndOfDoc);
    }
}
