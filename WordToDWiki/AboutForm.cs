using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordToDWiki
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
        }

        private void About_Load(object sender, EventArgs e)
        {
            rtxtAbout.Text += "Word To Doku Wiki v2.0" + Environment.NewLine;
            rtxtAbout.Text += "Internal libraries:" + Environment.NewLine;
            rtxtAbout.Text += "- MicroMWordLib v1.0" + Environment.NewLine;
            rtxtAbout.Text += "- LittleLyreLogger v1.0" + Environment.NewLine;
            rtxtAbout.Text += "- LittleImage v1.0" + Environment.NewLine;
            rtxtAbout.Text += "- DokuWikiFormatter v1.0" + Environment.NewLine;
            rtxtAbout.Text += Environment.NewLine;
            rtxtAbout.Text += "Software developer: Shamo Humbatli" + Environment.NewLine;
            rtxtAbout.Text += "E-mail: shamohumbatli@gmail.com" + Environment.NewLine;
            rtxtAbout.Text +=  Environment.NewLine;
            rtxtAbout.Text += "This is a free software. [2017 - 2018]" + Environment.NewLine;
            rtxtAbout.Text += "Note: if you find a bug or if you have any offer, please don't hesitate to email me." + Environment.NewLine;
        }
    }
}
