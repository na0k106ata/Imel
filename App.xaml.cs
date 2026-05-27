using System;
using System.Threading;
using System.Windows;

namespace Imel
{
    /// <summary>
    /// アプリケーションのエントリーポイント定義。
    /// WPFアプリケーションのライフサイクルと多重起動防止を管理します。
    /// </summary>
    public partial class App : Application
    {
        private const string SingleInstanceMutexName = @"Local\Imel_SingleInstance";
        private Mutex? _singleInstanceMutex;

        protected override void OnStartup(StartupEventArgs e)
        {
            _singleInstanceMutex = new Mutex(true, SingleInstanceMutexName, out bool createdNew);
            if (!createdNew)
            {
                _singleInstanceMutex.Dispose();
                _singleInstanceMutex = null;
                Shutdown();
                return;
            }

            base.OnStartup(e);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            if (_singleInstanceMutex != null)
            {
                _singleInstanceMutex.ReleaseMutex();
                _singleInstanceMutex.Dispose();
                _singleInstanceMutex = null;
            }

            base.OnExit(e);
        }
    }
}
