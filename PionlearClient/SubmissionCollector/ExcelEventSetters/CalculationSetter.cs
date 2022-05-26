using System;
using Microsoft.Office.Interop.Excel;

namespace SubmissionCollector.ExcelEventSetters
{
    internal class CalculationSetter : IDisposable
    {
        private static int _counter;
        private static XlCalculation _originalCalculation;
        public CalculationSetter()
        {
            _counter++;
            TryChangeState();
        }

        public void OnEnter()
        {
            _originalCalculation = Globals.ThisWorkbook.Application.Calculation;
            Globals.ThisWorkbook.Application.Calculation = XlCalculation.xlCalculationManual;
        }

        public void OnExit()
        {
            Globals.ThisWorkbook.Application.Calculation = _originalCalculation;
        }

        private void TryChangeState()
        {
            if (_counter < 0)
            {
                throw new InvalidOperationException("Excel calculation updating in bad state");
            }

            if (_counter == 0)
            {
                //Globals.ThisWorkbook.CurrentDispatcher.Invoke(() => OnExit());
                OnExit();
            }

            if (_counter > 0)
            {
                OnEnter();
            }
        }

        public void Dispose()
        {
            _counter--;
            TryChangeState();
        }
    }
}
