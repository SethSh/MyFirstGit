using System;
using PionlearClient;

namespace SubmissionCollector.ExcelEventSetters
{
    public class WorkbookUnprotector : IDisposable
    {
        private static int _counter;

        public WorkbookUnprotector()
        {
            _counter++;
            TryChangeState();
        }

        public void OnEnter()
        {
            Globals.ThisWorkbook.Unprotect(password: BexConstants.WorkbookPassword);
        }

        public void OnExit()
        {
            Globals.ThisWorkbook.Protect(password:BexConstants.WorkbookPassword);
        }

        private void TryChangeState()
        {
            if (_counter < 0)
            {
                throw new InvalidOperationException("Excel workbook protection in bad state");
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