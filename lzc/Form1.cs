using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace lzc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            notifyIcon1.Icon = SystemIcons.Application;
        }

        private void contextMenuStrip1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
