using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LoanCalculator.LoanCalculator
{
    public class MonthlyPayment
    {
        public int Id { get; set; }
        public decimal MonthlyInterestPayment { get; set; }
        public decimal LoanRemainder { get; set; }
        public decimal MonthlyLoanService { get; set; }
        public decimal PrincipalAmountRepayment { get; set; }
        public decimal TotalMonthlyPayment { get; set; }

        public MonthlyPayment(int id, decimal monthlyInterestPayment, 
            decimal loanRemainder, decimal monthlyLoanService, decimal paymentAmount)
        {
            this.Id = id;
            this.MonthlyInterestPayment = monthlyInterestPayment;
            this.LoanRemainder = loanRemainder;
            this.MonthlyLoanService = monthlyLoanService;
            this.PrincipalAmountRepayment = paymentAmount - this.MonthlyInterestPayment;
            this.TotalMonthlyPayment = this.MonthlyInterestPayment + this.PrincipalAmountRepayment + this.MonthlyLoanService;
        }
    }
}
