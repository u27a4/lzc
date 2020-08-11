using System;
using System.ComponentModel.Composition;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static lzc.NativeMethods;

namespace lzc.ext.lumos
{
    [PartCreationPolicy(CreationPolicy.NonShared)]
    class Lumos
    {
        private static IntPtr Hook;
        private static HOOKPROC Proc;
        private static Form Owner;

        // ライトを有効化するメニューを作成
        [MenuItem]
        static MenuItem CreateMenuItem()
        {
            return new MenuItem($"集中モードの有効化", OnMenuItemClicked)
            {
                Checked = (Owner != null),
            };
        }

        // メニューがクリックされた
        static void OnMenuItemClicked(object sender, EventArgs e)
        {
            if (sender is MenuItem menu)
            {
                if (Owner == null)
                {
                    // 有効化
                    if ((Hook = HookMouse()) != IntPtr.Zero)
                    {
                        Owner = new Form1();
                        Owner.Show();
                    }
                }
                else
                {
                    // 無効化
                    if (Hook != IntPtr.Zero)
                    {
                        UnhookWindowsHookEx(Hook);
                        Owner.Close();
                        Owner = null;
                    }
                }
            }
        }

        // マウスフック処理を登録
        static IntPtr HookMouse()
        {
            return SetWindowsHookEx(14, Proc = MouseHookProc, IntPtr.Zero, IntPtr.Zero); // WH_MOUSE_LL
        }

        // マウスフック処理
        static int MouseHookProc(int nCode, uint wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                if (wParam == 0x0200)   // WM_MOUSEMOVE
                {
                    if (Owner != null)
                    {
                        var mouse = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

                        Owner.Location = mouse.Point;
                    }
                }
            }

            return CallNextHookEx(Hook, nCode, wParam, lParam);
        }
    }

    class Form1 : Form
    {
        // フォームがロードされたときの処理
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            FormBorderStyle = FormBorderStyle.None;
            Size = Size.Empty;
            Move += Form_Move;
        }

        // フォームが表示されたときの処理
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            // ディスプレイの数だけ黒板を作成
            foreach (var screen in Screen.AllScreens)
            {
                var wa = screen.WorkingArea;
                var center = new Point(wa.X + wa.Width / 2, wa.Y + wa.Height / 2);

                var form = new Form();

                form.FormBorderStyle = FormBorderStyle.None;
                form.WindowState = FormWindowState.Maximized;
                form.StartPosition = FormStartPosition.Manual;
                form.Location = center;
                form.BackColor = Color.Black;
                form.Move += Form_Move;
                form.Shown += Form_Move;

                form.Show(this);
            }
        }

        // フォームが移動したときの処理
        private void Form_Move(object sender, EventArgs e)
        {
            foreach (var form in OwnedForms)
            {
                form.FormBorderStyle = FormBorderStyle.None;
                form.Visible = Screen.FromControl(form).DeviceName != Screen.FromControl(this).DeviceName;
            }
        }
    }
}
