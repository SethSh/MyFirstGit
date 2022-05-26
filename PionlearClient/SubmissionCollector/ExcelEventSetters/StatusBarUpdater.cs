using System;

namespace SubmissionCollector.ExcelEventSetters
{
    public class StatusBarUpdater : IDisposable
    {
        private static int _counter;
        private readonly string _status;

        public StatusBarUpdater(string status)
        {
            _counter++;
            _status = status;
            TryChangeState();
        }

        public void OnEnter()
        {
            Globals.ThisWorkbook.Application.StatusBar = _status;
        }

        public void OnExit()
        {

            try
            {
                Globals.ThisWorkbook.Application.StatusBar = false;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                //eat com exception
            }
        }

        private void TryChangeState()
        {
            if (_counter < 0)
            {
                throw new InvalidOperationException("Status bar updating event in bad state");
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
