using System;
using System.Windows;
using System.Windows.Input;

namespace LoanCalculator
{
    /// <summary>
    /// Расчет кредитной ставки
    /// </summary>
    public partial class MainWindow : Window
    {
        public ICalculator LoanCalculator { get; set; }

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

            decimal downPayment;
            int monthMultiplier = 1;
            bool tryDownPayment = decimal.TryParse(downPaymentTextBox.Text, out downPayment);
            if (!tryDownPayment)
            {
                downPayment = 0.0m;
            }
            if ((bool)yearsRadio.IsChecked)
            {
                monthMultiplier *= 12;
            }

            this.LoanCalculator = new Calculator(downPayment, Convert.ToDecimal(loanTextBox.Text),
                Convert.ToDecimal(annualAccrualsTextBox.Text),
                Convert.ToInt32(creditingPeriodTextBox.Text),
                monthMultiplier);
            txtSummary.Text += this.LoanCalculator.Calculate();
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (txtSummary.Text == "")
            {
                MessageBox.Show("Спочатку прорахуйте кредитну ставку.");
                return;
            }
            else
            {
                this.LoanCalculator.WordExport(System.Reflection.Missing.Value, "\\endofdoc");
            }
        }
    }
}
