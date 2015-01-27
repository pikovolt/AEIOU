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
    public partial class InputBox1 : Form
    {
        public InputBox1(String title)
        {
            InitializeComponent();

            this.Text = title;
            this.label1.Enabled = false;
            this.label2.Enabled = false;
            this.textBox1.Enabled = false;
            this.textBox2.Enabled = false;
            this.checkBox1.Enabled = false;
            this.label1.Text = "";
            this.label2.Text = "";
            this.checkBox1.Text = "";
        }

        public String LabelName1
        {
            set
            {
                this.label1.Text = value;
                this.label1.Enabled = true;
                this.textBox1.Enabled = true;
            }
        }

        public String LabelName2
        {
            set
            {
                this.label2.Text = value;
                this.label2.Enabled = true;
                this.textBox2.Enabled = true;
            }
        }

        public String CheckName1
        {
            set
            {
                this.checkBox1.Text = value;
                this.checkBox1.Enabled = true;
            }
        }

        public String Value1
        {
            get
            {
                return this.textBox1.Text;
            }
            set
            {
                this.textBox1.Text = value;
            }
        }

        public String Value2
        {
            get
            {
                return this.textBox2.Text;
            }
            set
            {
                this.textBox2.Text = value;
            }
        }

        public bool CheckValue1
        {
            get
            {
                return this.checkBox1.Checked;
            }
            set
            {
                this.checkBox1.Checked = value;
            }
        }
    }
}
