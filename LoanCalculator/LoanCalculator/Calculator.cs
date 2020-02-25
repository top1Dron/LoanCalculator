using LoanCalculator.LoanCalculator;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;

namespace LoanCalculator
{
    public class Calculator : ICalculator
    {
        public decimal LoanAmount { get; set; }
        public decimal DownPayment { get; set; }
        public decimal AnnualInterestRate { get; set; }
        public int LoanPeriod { get; set; }
        public decimal SetAndComService { get; set; }
        public decimal LoanProcessing { get; set; }
        public decimal NotarialServices { get; set; }
        public decimal LoanService { get; set; }
        public decimal OutpostInsurance { get; set; }
        public decimal PropertyValuationService { get; set; }
        public decimal TotalAmountRepaid { get; set; }
        public decimal TotalInterest { get; set; }
        public String Summary { get; set; }

        public List<MonthlyPayment> MonthlyPayments { get; }

        public Calculator(decimal downPayment, decimal loanAmount, decimal annualInterestRate, 
            int loanPeriod, int monthMultiplier, decimal setAndComService,
            decimal loanProcessing, decimal notarialServices, decimal loanService,
            decimal outpostInsurance, decimal propertyValuationService)
        {
            this.DownPayment = downPayment;
            this.LoanAmount = loanAmount - this.DownPayment;
            this.AnnualInterestRate = annualInterestRate;
            this.LoanPeriod = loanPeriod;
            this.LoanPeriod *= monthMultiplier;
            this.SetAndComService = setAndComService / 100 * this.LoanAmount;
            this.LoanProcessing = loanProcessing / 100 * this.LoanAmount;
            this.NotarialServices = notarialServices;
            this.LoanService = loanService / 100;
            this.OutpostInsurance = outpostInsurance;
            this.PropertyValuationService = propertyValuationService;
            this.MonthlyPayments = new List<MonthlyPayment>(150);
        }

        public String Calculate()
        {
            this.Summary = "";
            var interestRate = (this.AnnualInterestRate / 100) / 12;

            var paymentAmount = (interestRate * this.LoanAmount);
            var divPaymentAmount = 1 - (Math.Pow(1 + (Convert.ToDouble(interestRate)), -(Convert.ToDouble(this.LoanPeriod))));
            paymentAmount /= Convert.ToDecimal(divPaymentAmount);
            paymentAmount = Math.Round(paymentAmount, 2);


            this.TotalAmountRepaid = paymentAmount * LoanPeriod + this.DownPayment + this.SetAndComService + 
                this.LoanProcessing + this.NotarialServices + this.PropertyValuationService + this.OutpostInsurance;

            var remainder = this.LoanAmount;
            this.TotalInterest = 0m;
            var montlyInterestCoefficient = 12m * (365 * 4 + 1) / (48 * 365.25m);  // average days per month and days per year are presented here
            for (var j = 1; j <= this.LoanPeriod; j++)
            {
                var monthlyInterestPayment = remainder * interestRate * montlyInterestCoefficient;
                var monthlyLoanService = this.LoanService * remainder;
                this.TotalAmountRepaid += monthlyLoanService;
                remainder += monthlyInterestPayment - paymentAmount;
                this.MonthlyPayments.Add(new MonthlyPayment(j, Math.Round(monthlyInterestPayment, 2), Math.Round(remainder, 2), Math.Round(monthlyLoanService, 2), paymentAmount));

                //Debug.WriteLine(Math.Round(monthlyInterestPayment, 2) + " " + Math.Round(remainder, 2) + " " + Math.Round(monthlyLoanService, 2));
                this.TotalInterest += monthlyInterestPayment;
            }


            this.Summary = "Загальна кількість платежів - " + LoanPeriod.ToString() +
                "\r\nРозрахунково-касове обслуговування - " + this.SetAndComService.ToString("N") +
                "\r\nКомісія за надання кредиту - " + this.LoanProcessing.ToString("N") +
                "\r\nПослуги нотаріуса - " + this.NotarialServices.ToString("N") +
                "\r\nПослуги з оцінки майна - " + this.PropertyValuationService.ToString("N") +
                "\r\nПовна сума виплати становить - " + TotalAmountRepaid.ToString("N") + "." +
                "\r\nЗагальна сума відсотків, сплачених за період позики, становить - " + TotalInterest.ToString("N");
            return this.Summary;
        }

        public void WordExport(object oMissing, object oEndOfDoc)
        {
            var dataTable = new System.Data.DataTable();
            dataTable.Columns.Add("Номер місяця", typeof(int));
            dataTable.Columns.Add("Залишок позики", typeof(decimal));
            dataTable.Columns.Add("Сума за розрахунковий період", typeof(decimal));
            dataTable.Columns.Add("Погашення основної суми кредиту", typeof(decimal));
            dataTable.Columns.Add("Процент за вик. кредиту", typeof(decimal));
            dataTable.Columns.Add("Управління кредитом", typeof(decimal));
            //dataTable.Rows.Add("Номер місяця", "Залишок позики", "Сума за розрахунковий період",
            //    "Погашення основної суми кредиту", "Процент за використання кредиту", "Управління кредитом");
            foreach(MonthlyPayment monthlyPayment in MonthlyPayments)
            {
                dataTable.Rows.Add(monthlyPayment.Id, monthlyPayment.LoanRemainder, monthlyPayment.TotalMonthlyPayment,
                    monthlyPayment.PrincipalAmountRepayment, monthlyPayment.MonthlyInterestPayment, monthlyPayment.MonthlyLoanService);
            }

            //Start Word and create a new document.
            _Application wordApplication = new Application();
            _Document wordDocument = null;
            wordApplication.Visible = true;
            try
            {
                wordDocument = wordApplication.Documents.Add(ref oMissing,
                ref oMissing, ref oMissing, ref oMissing);
            }
            catch (Exception)
            {
                if (wordDocument != null)
                {
                    wordDocument.Close(false);
                    wordDocument = null;
                }
                wordApplication.Quit();
                wordApplication = null;
                return;
            }

            wordApplication.Selection.Find.Execute("%метка%");
            Range wordRange = wordApplication.Selection.Range;

            var wordTable = wordDocument.Tables.Add(wordRange,
                dataTable.Rows.Count+1, dataTable.Columns.Count);
            wordTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleDouble;
            wordTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleDouble;

            for (int columnNumber = 0; columnNumber < dataTable.Columns.Count; columnNumber++)
            {
                wordTable.Cell(1, columnNumber + 1).Range.Text = dataTable.Columns[columnNumber].ColumnName;
            }

            for (var j = 1; j < dataTable.Rows.Count+1; j++)
            {
                for (var k = 0; k < dataTable.Columns.Count; k++)
                {
                    wordTable.Cell(j + 1, k + 1).Range.Text = dataTable.Rows[j-1][k].ToString();
                }
            }

            //Insert a paragraph at the beginning of the document.
            Paragraph oPara1;
            oPara1 = wordDocument.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Розмір кредиту - " + this.LoanAmount.ToString() + " грн.\n" +
                "Перший внесок - " + this.DownPayment.ToString() + " грн.\n" +
                "Річні нарахування - " + this.AnnualInterestRate.ToString() + " %\n" +
                this.Summary;
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();
        }

        public static decimal Pow(decimal b, int e)
        {
            decimal multiple = b;
            for (var i = 0; i < e; i++)
            {
                multiple *= b;
            }
            return multiple;
        }
    }
}
