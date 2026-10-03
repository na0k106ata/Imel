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

        /// <summary>アプリの終了処理（トレイからの終了・サインアウト）が始まっているか。</summary>
        internal static bool IsExiting { get; private set; }

        internal static void RequestExit()
        {
            // ウィンドウを閉じる順序に関係なく、終了中であることを先に共有する。
            IsExiting = true;
            Current.Shutdown();
        }

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

            StartupRegistration.RefreshIfMoved();
            base.OnStartup(e);
        }

        protected override void OnSessionEnding(SessionEndingCancelEventArgs e)
        {
            IsExiting = true;
            base.OnSessionEnding(e);
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
