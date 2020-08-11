using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using static lzc.NativeMethods;

namespace lzc
{
    static class Program
    {
        private static Form1 Form1;
        private static IntPtr Hook;
        private static HOOKPROC Proc;

        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            if ((Hook = HookKeyboard()) != IntPtr.Zero)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(Form1 = new Form1());

                UnhookWindowsHookEx(Hook);
            }
            else
            {
                MessageBox.Show("An error has occurred trying to hook keyboard", "lzc");
            }
        }

        // キーボードフック処理を登録
        private static IntPtr HookKeyboard()
        {
            return SetWindowsHookEx(13, Proc = KeyboardHookProc, IntPtr.Zero, IntPtr.Zero); // WH_KEYBOARD_LL
        }

        // キーボードフック処理
        private static int KeyboardHookProc(int nCode, uint wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                if (wParam == 0x0100)   // WM_KEYDOWN
                {
                    var kb = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                    var primary1 = GetAsyncKeyState(0x5B) & 0x8000; // VK_LWIN
                    var primary2 = GetAsyncKeyState(0x5C) & 0x8000; // VK_RWIN,
                    var secondary = (kb.vkCode == 0x5A);            // VK_Z

                    if ((primary1 | primary2) != 0 & secondary)
                    {
                        var menu = BuildMenu(@"./ext");
                        var x = Cursor.Position.X;
                        var y = Cursor.Position.Y;

                        SetForegroundWindow(Form1.Handle);
                        TrackPopupMenu(menu.Handle, 0, x, y, 0, Form1.Handle, IntPtr.Zero);
                    }
                }
            }

            return CallNextHookEx(Hook, nCode, wParam, lParam);
        }

        // メニューの組み立て
        private static ContextMenu BuildMenu(string rootDir)
        {
            Directory.CreateDirectory(rootDir);

            var assm = new AssemblyCatalog(Assembly.GetExecutingAssembly());
            var exts = new DirectoryCatalog(rootDir);
            var aggr = new AggregateCatalog(assm, exts);
            var container = new CompositionContainer(aggr);

            return new ContextMenu(container.GetExportedValues<Func<MenuItem>>()
                .Select(TryCreateMenu)
                .Where(item => item != null)
                .ToArray());
        }

        // メニューの生成
        private static MenuItem TryCreateMenu(Func<MenuItem> callback)
        {
            try
            {
                return callback();
            }
            catch (Exception e)
            {
                return new MenuItem($"failed to add menu: {e.Source}");
            }
        }
    }
}
