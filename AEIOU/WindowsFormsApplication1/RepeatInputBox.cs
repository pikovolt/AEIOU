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
    public sealed class RepeatInputValues
    {
        public RepeatInputValues(string start, string end, string rowInterval,
            string loop, string skip, string insert)
        {
            Start = start;
            End = end;
            RowInterval = rowInterval;
            Loop = loop;
            Skip = skip;
            Insert = insert;
        }

        public string Start { get; private set; }
        public string End { get; private set; }
        public string RowInterval { get; private set; }
        public string Loop { get; private set; }
        public string Skip { get; private set; }
        public string Insert { get; private set; }
    }

    public partial class RepeatInputBox : Form
    {
        public event Action<RepeatInputValues> OnRepeatInput;
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
        private void buttonRepeat_Click(object sender, EventArgs e)
        {
            Action<RepeatInputValues> handler = OnRepeatInput;
            if (handler != null)
                handler(new RepeatInputValues(Value1, Value2, Value3, Value4, Value5, Value6));
        }
    }
}
