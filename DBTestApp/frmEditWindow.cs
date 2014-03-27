using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace DBTestApp
{
    public partial class frmEditWindow : Form
    {
        public frmEditWindow()
        {
            InitializeComponent();
        }

        private void frmEditWindow_Load(object sender, EventArgs e)
        {
         //Nothing needed here yet   
        }

        
        /// <summary>
        /// User has clicked СОХРАНИТЬ
        /// </summary>
        /// <param name="sender">button1</param>
        /// <param name="e">event args</param>
        private void button1_Click(object sender, EventArgs e)
        {
            frmEditWindow.ActiveForm.Close();
            Form1._save = true;
        }

        /// <summary>
        /// User has clicked on a link
        /// </summary>
        /// <param name="sender">richTextBox1</param>
        /// <param name="e">event args</param>
        private void richTextBox1_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            //
            // extracting URL from event and navigate using windows default web browser
            //
            string target = e.LinkText as string;
            try
            {
                if (null != target)
                {
                    System.Diagnostics.Process.Start(target);
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }

        }
        /// <summary>
        /// User has clicked on richtextbox1 field and we help him with a link :)
        /// </summary>
        /// <param name="sender">richTextBox1 Control</param>
        /// <param name="e">event args</param>
        private void richTextBox1_Click(object sender, EventArgs e)
        {
           
            richTextBox1.Text = "http://";
            richTextBox1.SelectionStart = richTextBox1.TextLength;
        }

        
    }
}
