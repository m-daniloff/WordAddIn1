using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Esquire.Common.ViewModels;
using System.Threading;
using System.Windows.Threading;

namespace Esquire.Common.WPF.ViewModels
{
    public class RunThreadProgressBar : ViewModelBase
    {
        private Thread _runThread;
        private WPF.RunThreadProgressBar _window;
        private string _messageText;

        public RunThreadProgressBar(Thread thread, string messageText)
        {
            _runThread = thread;
            _window = new WPF.RunThreadProgressBar();
            _window.DataContext = this;
            Message = messageText;
        }

        public string Message
        {
            get { return _messageText; }
            set
            {
                _messageText = value;
                OnPropertyChanged("Message");
            }
        }

        public void Start()
        {
            var mainDispatcher = Dispatcher.FromThread(Thread.CurrentThread);
            (new Thread(() =>
            {
                _runThread.Start();
                _runThread.Join();
                if (_window.IsVisible)
                {
                    try
                    {
                        mainDispatcher.BeginInvoke(new Action(_window.Close));
                    }
                    catch { }
                }
            })).Start();

            _window.ShowDialog();
        }
    }
}
