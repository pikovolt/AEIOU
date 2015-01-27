using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AEIOU
{
    public partial class RepeatInputBox : Form
    {
        public RepeatInputBox()
        {
            InitializeComponent();
        }

        public String Value1
        {
            get { return textBox1.Text; }
        }
        public String Value2
        {
            get { return textBox2.Text; }
        }
        public String Value3
        {
            get { return textBox3.Text; }
        }
        public String Value4
        {
            get { return textBox4.Text; }
        }
        public String Value5
        {
            get { return textBox5.Text; }
        }
        public String Value6
        {
            get { return textBox6.Text; }
        }
    }
}
