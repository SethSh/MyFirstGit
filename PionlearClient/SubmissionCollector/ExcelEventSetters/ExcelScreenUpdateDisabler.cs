using System;

namespace SubmissionCollector.ExcelEventSetters
{
    public class ExcelScreenUpdateDisabler : IDisposable
    {
        private static int _counter;

        public ExcelScreenUpdateDisabler()
        {
            _counter++;
            TryChangeState();
        }

        public void OnEnter()
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
        }

        public void OnExit()
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }

        private void TryChangeState()
        {
            if (_counter < 0)
            {
                throw new InvalidOperationException("Screen updating event in bad state");
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
