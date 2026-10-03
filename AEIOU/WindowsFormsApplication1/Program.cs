using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

namespace AEIOU
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 起動時に少し待機
            Thread.Sleep(1000); // 1秒待機

            Application.Run(new Form1());
        }
    }
}
