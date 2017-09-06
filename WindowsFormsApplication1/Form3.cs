using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        public void display(String path)
        {

           // axShockwaveFlash1.Movie = Application.StartupPath + @path;
            axShockwaveFlash1.Movie =  @path;
            axShockwaveFlash1.Playing = true;
        
        }

    }
}
