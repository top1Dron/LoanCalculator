using Microsoft.Office.Interop.Word;
using System;

namespace LoanCalculator
{
    public class Calculator : ICalculator
    {
        public decimal LoanAmount { get; set; }
        public decimal DownPayment { get; set; }
        public decimal AnnualInterestRate { get; set; }
        public int LoanPeriod { get; set; }
        public decimal TotalAmountRepaid { get; set; }
        public decimal TotalInterest { get; set; }

        public String Summary { get; set; }

        public Calculator(decimal downPayment, decimal loanAmount, decimal annualInterestRate, int loanPeriod, int monthMultiplier)
        {
            this.DownPayment = downPayment;
            this.LoanAmount = loanAmount - this.DownPayment;
            this.AnnualInterestRate = annualInterestRate;
            this.LoanPeriod = loanPeriod;
            this.LoanPeriod *= monthMultiplier;
        }

        public String Calculate()
        {
            this.Summary = "";
            decimal interestRate = (this.AnnualInterestRate / 100) / 12;

            decimal paymentAmount = (interestRate * this.LoanAmount);
            double divPaymentAmount = 1 - (Math.Pow(1 + (Convert.ToDouble(interestRate)), -(Convert.ToDouble(this.LoanPeriod))));
            paymentAmount = paymentAmount / Convert.ToDecimal(divPaymentAmount);
            paymentAmount = Math.Round(paymentAmount, 2);

            this.TotalAmountRepaid = paymentAmount * LoanPeriod + this.DownPayment;
            this.TotalInterest = this.TotalAmountRepaid - this.LoanAmount;

            this.Summary = "Загальна кількість платежів - " + (LoanPeriod).ToString() + ".\r\nЩомісячний платіж становить - " + paymentAmount.ToString("N") +
                "\r\nПовна сума виплати становить - " + TotalAmountRepaid.ToString("N") + "." +
                "\r\nЗагальна сума відсотків, сплачених за період позики, становить - " + TotalInterest.ToString("N");
            return this.Summary;
        }

        public void WordExport(object oMissing, object oEndOfDoc)
        {
            //Start Word and create a new document.
            _Application oWord;
            _Document oDoc;
            oWord = new Application
            {
                Visible = true
            };
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);

            //Insert a paragraph at the beginning of the document.
            Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Розмір кредиту - " + this.LoanAmount.ToString() + " грн.\n" +
                "Перший внесок - " + this.DownPayment.ToString() + " грн.\n" +
                "Річні нарахування - " + this.AnnualInterestRate.ToString() + " %\n" +
                this.Summary;
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();
        }
    }
}
