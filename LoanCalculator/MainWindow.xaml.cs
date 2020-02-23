using Microsoft.Office.Interop.Word;
using System;
using System.Windows;
using System.Windows.Input;

namespace LoanCalculator
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private decimal loanAmount;
        private decimal downPayment;
        private decimal annualInterestRate;
        private int loanPeriod;
        private decimal totalAmountRepaid;
        private decimal totalInterest;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void LoanTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789 ,".IndexOf(e.Text) < 0;
        }

        private void AnnualAccrualsTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789 ,".IndexOf(e.Text) < 0;
        }

        private void CreditingPeriodTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789 ,".IndexOf(e.Text) < 0;
        }

        private void BtnCalculate_Click(object sender, RoutedEventArgs e)
        {
            CalculateLoanSchedule();
        }

        private void CalculateLoanSchedule()
        {
            txtSummary.Text = "";
            if (loanTextBox.Text == "")
            {
                MessageBox.Show("Поле \"Розмір кредиту\" не може бути порожнім", "Помилка");
                return;
            }
            else if (creditingPeriodTextBox.Text == "")
            {
                MessageBox.Show("Поле \"Період кредитування\" не може бути порожнім", "Помилка");
                return;
            }
            else if (annualAccrualsTextBox.Text == "")
            {
                MessageBox.Show("Поле \"Річні нарахування\" не може бути порожнім", "Помилка");
                return;
            }

            bool tryDownPayment = decimal.TryParse(downPaymentTextBox.Text, out downPayment);
            if (!tryDownPayment)
            {
                downPayment = 0.0m;
            }
            loanAmount = Convert.ToDecimal(loanTextBox.Text) - downPayment;
            annualInterestRate = Convert.ToDecimal(annualAccrualsTextBox.Text);
            loanPeriod = Convert.ToInt32(creditingPeriodTextBox.Text);
            if ((bool)yearsRadio.IsChecked)
            {
                loanPeriod *= 12;
            }


            decimal interestRate = (annualInterestRate / 100) / 12;

            decimal paymentAmount = (interestRate * loanAmount);
            double divPaymentAmount = 1 - (Math.Pow(1 + (Convert.ToDouble(interestRate)), -(Convert.ToDouble(loanPeriod))));
            paymentAmount = paymentAmount / Convert.ToDecimal(divPaymentAmount);
            paymentAmount = Math.Round(paymentAmount, 2);

            totalAmountRepaid = paymentAmount * 12 + downPayment;
            totalInterest = totalAmountRepaid - loanAmount;

            txtSummary.Text += "Загальна кількість платежів - " + (loanPeriod).ToString() + ".\r\nЩомісячний платіж становить - " + paymentAmount.ToString("N") +
                "\r\nПовна сума виплати становить - " + totalAmountRepaid.ToString("N") + "." +
                "\r\nЗагальна сума відсотків, сплачених за період позики, становить - " + totalInterest.ToString("N");
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (txtSummary.Text == "")
            {
                MessageBox.Show("Спочатку прорахуйте кредитну ставку.");
                return;
            }

            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            _Application oWord;
            _Document oDoc;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);

            //Insert a paragraph at the beginning of the document.
            Microsoft.Office.Interop.Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Розмір кредиту - " + this.loanAmount.ToString() + " грн.\n" +
                "Перший внесок - " + this.downPayment.ToString() + " грн.\n" +
                "Річні нарахування - " + this.annualInterestRate.ToString() + " %\n" +
                txtSummary.Text;
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();
        }
    }
}
