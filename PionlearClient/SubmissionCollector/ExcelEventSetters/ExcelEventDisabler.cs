using System;

namespace SubmissionCollector.ExcelEventSetters
{
    public class ExcelEventDisabler : IDisposable
    {
        private static int _counter;

        public ExcelEventDisabler()
        {
            _counter++;
            TryChangeState();
        }

        public void OnEnter()
        {
            Globals.ThisWorkbook.Application.EnableEvents = false;
        }

        public void OnExit()
        {
            Globals.ThisWorkbook.Application.EnableEvents = true;
        }

        private void TryChangeState()
        {
            if (_counter < 0)
            {
                throw new InvalidOperationException("Enable events updating event in bad state");
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
