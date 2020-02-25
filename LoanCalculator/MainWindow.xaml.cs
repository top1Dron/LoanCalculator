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

        private void SetAndComServiceTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789 ,".IndexOf(e.Text) < 0;
        }

        private void LoanProcessingTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789 ,".IndexOf(e.Text) < 0;
        }

        private void NotarialServicesTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789 ,".IndexOf(e.Text) < 0;
        }

        private void OutpostCostTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789 ,".IndexOf(e.Text) < 0;
        }

        private void LoanServiceTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789 ,".IndexOf(e.Text) < 0;
        }

        private void OutpostInsuranceTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = "0123456789 ,".IndexOf(e.Text) < 0;
        }

        private void PropertyValuationServiceTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
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
            if (downPaymentTextBox.Text == "")
            {
                downPaymentTextBox.Text = "0";
            }
            if (setAndComServiceTextBox.Text == "")
            {
                setAndComServiceTextBox.Text = "0";
            }
            if (loanProcessingTextBox.Text == "")
            {
                loanProcessingTextBox.Text = "0";
            }
            if (notarialServicesTextBox.Text == "")
            {
                notarialServicesTextBox.Text = "0";
            }
            if (loanServiceTextBox.Text == "")
            {
                loanServiceTextBox.Text = "0";
            }
            if (outpostCostTextBox.Text == "")
            {
                outpostCostTextBox.Text = "0";
            }
            if (outpostInsuranceTextBox.Text == "")
            {
                outpostInsuranceTextBox.Text = "0";
            }
            if (propertyValuationServiceTextBox.Text == "")
            {
                propertyValuationServiceTextBox.Text = "0";
            }
            if(Convert.ToDecimal(downPaymentTextBox.Text) > Convert.ToDecimal(loanTextBox.Text))
            {
                MessageBox.Show("Поле \"Початковий внесок\" не може мати значення більше, ніж поле \"Розмір кредиту\"", "Помилка");
                return;
            }

            int monthMultiplier = 1;
            if ((bool)yearsRadio.IsChecked)
            {
                monthMultiplier *= 12;
            }

            decimal outpostInsurance = Convert.ToDecimal(outpostCostTextBox.Text) * Convert.ToDecimal(outpostInsuranceTextBox.Text) / 100;

            this.LoanCalculator = new Calculator(Convert.ToDecimal(downPaymentTextBox.Text), Convert.ToDecimal(loanTextBox.Text),
                Convert.ToDecimal(annualAccrualsTextBox.Text),
                Convert.ToInt32(creditingPeriodTextBox.Text),
                monthMultiplier, Convert.ToDecimal(setAndComServiceTextBox.Text),
                Convert.ToDecimal(loanProcessingTextBox.Text), Convert.ToDecimal(notarialServicesTextBox.Text),
                Convert.ToDecimal(loanServiceTextBox.Text), outpostInsurance, Convert.ToDecimal(propertyValuationServiceTextBox.Text));
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
