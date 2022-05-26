using System;
using Microsoft.Office.Interop.Excel;

namespace SubmissionCollector.ExcelEventSetters
{
    public class CursorToWait : IDisposable
    {
        private static int _counter;

        public CursorToWait()
        {
            _counter++;
            TryChangeState();
        }
        public void OnEnter()
        {
            Globals.ThisWorkbook.Application.Cursor = XlMousePointer.xlWait;
        }

        public void OnExit()
        {
            Globals.ThisWorkbook.Application.Cursor = XlMousePointer.xlDefault;
        }

        private void TryChangeState()
        {
            if (_counter < 0)
            {
                throw new InvalidOperationException("Excel cursor in bad state");
            }

            if (_counter == 0)
            {
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