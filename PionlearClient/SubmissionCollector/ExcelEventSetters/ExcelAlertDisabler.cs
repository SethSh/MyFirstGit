using System;

namespace SubmissionCollector.ExcelEventSetters
{
    public class ExcelAlertDisabler : IDisposable
    {
        private static int _counter;

        public ExcelAlertDisabler()
        {
            _counter++;
            TryChangeState();
        }

        public void OnEnter()
        {
            Globals.ThisWorkbook.Application.DisplayAlerts = false;
        }

        public void OnExit()
        {
            Globals.ThisWorkbook.Application.DisplayAlerts = true;
        }

        private void TryChangeState()
        {
            if (_counter < 0)
            {
                throw new InvalidOperationException("Excel alerts updating event in bad state");
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
