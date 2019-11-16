using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace scaner
{
    
    public partial class Form2 : Form
    {
        
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            label1.BackColor = Color.Transparent;
            /*pictureBox1.BackColor = Color.Transparent;
            Graphics g = Graphics.FromHwnd(Handle);
            Pen p = new Pen(Color.Red);
            g.DrawLine(p, new Point(0, 0), new Point(10, 0));*/
            


        }
    }
}
