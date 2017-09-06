using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Diagnostics;





namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        
        public TimeSpan ts;
        public string[] comp,comp1;
        public string[] last,last1;
        public double[] per,per1;
        public string path,path_img,path_flash,colour1,picflag="not",pic;
        public int count = 0, i = 1, speed =0,sp=1,checker=1,tmp=1,cnt=1,refresh,u=0;
        public Form3 form3 = new Form3();
        public Stopwatch stopWatch = new Stopwatch();
        public PointF vector = new PointF();
        public Panel temp = new Panel();

        public Color colour = System.Drawing.Color.FromName("White");
        public Color fontcolour = System.Drawing.Color.FromName("Yellow");
   



        public Form2()
        {
            InitializeComponent();
            
        }

        private void Form2_Load(object sender, EventArgs e)
        {
        }
        public void start(string f,string pic2)
        { pic=pic2;
       
        if (pic != null)
        {
            panel12.Location = new Point(0, 0);
            
            path_img = pic; panel12.BackgroundImage = new Bitmap(@path_img);
            var pic1 = new Bitmap(panel12.BackgroundImage, new Size(panel12.Width, panel12.Height));
            panel12.BackgroundImage = pic1;
            panel12.Visible = true;
            picflag = null;
            path = f; excelprocess();
           // Thread.Sleep(150);
          //  MessageBox.Show("okeee");
           // panel12.Visible = false;
            u = 1; 
            panel12.Location = new Point(panel12.Location.X+1, panel12.Location.Y );
           timer1.Start();
        }
        else
        { panel12.Location = new Point(panel12.Location.X+1, panel12.Location.Y); path = f; excelprocess(); }
        }

        public string excelfile()
        {
            openFileDialog1.Filter = "Excel Files (XLS,XLSX)|*.xls;*.xlsx";
            DialogResult dlgResult = openFileDialog1.ShowDialog();
            if (dlgResult.Equals(DialogResult.OK))
            {
                path = openFileDialog1.FileName;
                return openFileDialog1.FileName;
            }
            return null;
        }
        public string img_file()
        {
            openFileDialog2.Filter = "Image Files (JPG,PNG,GIF,Exif,TIFF,RIF, BMP PPM, PGM, PBM,PNM,WEBP)|*.JPG;*.PNG;*.GIF;*.Exif;*.TIFF;*.RIF;*.BMP;*.PPM;*.PGM;*.PBM;*.PNM;*.WEBP";
            
            DialogResult dlgResult = openFileDialog2.ShowDialog();
            if (dlgResult.Equals(DialogResult.OK))
            {
                path_img = openFileDialog2.FileName;
             pic = openFileDialog2.FileName;
             MessageBox.Show("Image LoadedSucces fully");
                return openFileDialog2.FileName;
            }
            return null;
            
        }

        public string flash_file()
        {
            openFileDialog3.Filter = "Video Files (SWF)|*.SWF";
            DialogResult dlgResult = openFileDialog3.ShowDialog();
            if (dlgResult.Equals(DialogResult.OK))
            {
                path_flash = openFileDialog3.FileName;
                return openFileDialog3.FileName;
            }
            return null;
        }

        public void Color(Color s)
        {
           
            colour = s;

        }
        public void fontColor(Color s1)
        {

            fontcolour = s1;

        }
        public void speeds(int s)
        {
            if (s == 1)
            { sp = 1; this.timer1.Interval = 30; }
            else if (s == 2)
            { sp = 1; this.timer1.Interval = 15; }
            else
            { sp = 1; this.timer1.Interval = 1; }

        }
        public void tymstopper()
        {
           
            timer1.Stop();
            
        }
        public void starttimer()
        {
            this.BackgroundImage = null;
            form3.Hide();
            timer1.Start();
           
        }
        public void image()
        {
            panel1.Visible = false; panel2.Visible = false; panel3.Visible = false;
            panel4.Visible = false; panel5.Visible = false; panel6.Visible = false;
            panel8.Visible = false; panel10.Visible = false;
            panel7.Visible = false; panel9.Visible = false; panel11.Visible = false;
            form3.Hide();
            //Size s= new Size(227, 171);
             
            //Bitmap b = new Bitmap(160,416);
            //Bitmap objBitmap = new Bitmap(@path,s);
            this.BackgroundImage = new Bitmap(@path_img);
            var pic = new Bitmap(this.BackgroundImage, new Size(this.Width, this.Height));
            this.BackgroundImage = pic;    
        }

        public void rmpic()
        {
            if (path_img != null)
            {
                path_img = null;
                System.IO.File.WriteAllText(@".\file2.txt", string.Empty);
                pic = null;
                MessageBox.Show("Image Remove Successfully...");
            }
            else
                MessageBox.Show("Currently Image is not Loaded...");
        }

        public void flash()
        {
            panel1.Visible = false; panel2.Visible = false; panel3.Visible = false;
            panel4.Visible = false; panel5.Visible = false; panel6.Visible = false;
             panel8.Visible = false; panel10.Visible = false;
            panel7.Visible = false; panel9.Visible = false; panel11.Visible = false;
            this.BackgroundImage = null;
             form3.Tag = this;
            form3.Show(this);
            form3.display(path_flash);
            
            
            // Hide();
            //axShockwaveFlash1.Movie = Application.StartupPath + @path_flash;
        }

        public void excelprocess()
        {
        
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range; 
            string str;
            int rCnt = 0;
            int cCnt = 0;
            count = 0;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            comp = new string[range.Rows.Count + 1];
            last = new string[range.Rows.Count + 1];
            per = new double[range.Rows.Count + 1];
            comp1 = new string[range.Rows.Count + 1];
            last1 = new string[range.Rows.Count + 1];
            per1 = new double[range.Rows.Count + 1];

            refresh = range.Rows.Count + 8;
            for (rCnt = 1; rCnt < range.Rows.Count; rCnt++)
            {
                count++;
                for (cCnt = 1; cCnt <= 7; cCnt++)
                {
                    if (cCnt == 2 || cCnt == 3 || cCnt == 7)
                    {
                        try
                        {
                            str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                        }
                        catch (System.Exception excep)
                        { str = ((range.Cells[rCnt, cCnt] as Excel.Range).Value2).ToString("0.00"); }

                        // try
                        //{
                        if (cCnt == 2)
                        {
                            if (checker == 0)
                                comp1[count] = str;
                            else
                                comp[count] = str;
                        }
                        else if (cCnt == 3)
                        {
                            if (checker == 0)
                                last1[count] = str;
                            else
                                last[count] = str;
                        }
                        else if (cCnt == 7)
                        {
                            if (checker == 0)
                                per1[count] = Convert.ToDouble(str);
                            else
                                per[count] = Convert.ToDouble(str);
                        }
                        //}
                        // catch (System.Exception excep)
                        // { MessageBox.Show("please enter Valid values in columns "); }
                        // MessageBox.Show(str);

                    }
                }
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            if (checker == 0)
            {
                checker = 1; //stopWatch.Restart();
               // ts = TimeSpan.Zero; //timer1.Start();
            }
            else
            panelstart();


        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


        private void panelstart()
        {
            this.BackgroundImage = null;
            form3.Hide();
            Control[] controls;
            Double sign;
          
        /*   this.panel1.Location = new System.Drawing.Point(0,413);
            this.panel2.Location = new System.Drawing.Point(0, 1);
            this.panel3.Location = new System.Drawing.Point(0, 42);
            this.panel4.Location = new System.Drawing.Point(0, 83);
            this.panel5.Location = new System.Drawing.Point(0, 124);
            this.panel6.Location = new System.Drawing.Point(0, 165);
            this.panel2.Location = new System.Drawing.Point(0, 206);
            this.panel3.Location = new System.Drawing.Point(0, 247);
            this.panel4.Location = new System.Drawing.Point(0, 288);
            this.panel5.Location = new System.Drawing.Point(0, 329);
            this.panel6.Location = new System.Drawing.Point(0, 369);*/
            panel11.Visible = true;
            i = 1;
           // for (int j = 1; j <= 11; j++)
           //{
               int j=5;
                if (i > count)
                    i = 1;
                for (int k = 1; k <= 5; k++)
                {

                    //  MessageBox.Show("Eokeeeeeeeeeeeeeee");
                    controls = this.Controls.Find("Label" + j.ToString() + "_" + k, true);
                    if (controls.Length == 1) // 0 means not found, more - there are several controls with the same name
                    {
                        Label control = controls[0] as Label;
                        if (control != null)
                        {
                       

                       if (k == 1)
                       {  
                           control.Text = comp[i];
                          control.ForeColor = colour;
                       }
                       else if (k == 2)
                       {
                           control.Text = last[i];
                           control.ForeColor = fontcolour;
                       }
                       else if (k == 3)
                       {

                           if (per[i] > 0)
                           { control.Image = global::WindowsFormsApplication1.Properties.Resources.g; }
                           else if (per[i] < 0)
                               control.Image = global::WindowsFormsApplication1.Properties.Resources.r;
                           else
                               control.Image = global::WindowsFormsApplication1.Properties.Resources.y;

                       }
                       else if (k == 4)
                       {
                           control.ForeColor = fontcolour;
                          
                           
                           if (per[i] < 0)
                               sign = -1;
                           else
                               sign = 1;
                           control.Text = (per[i] * sign).ToString("0.00") + "%";
                       }
                       else
                       {
                               control.BackColor = colour;
                       }
                            

                            //   MessageBox.Show("arunnn");

                        }
                    }
                }


                i++;
               
            //}
            stopWatch.Restart();
            ts = TimeSpan.Zero;
          
           if(picflag!=null)
               timer1.Start();
           
           
        }


        public void refreshfile()
        { comp = null; last = null; per = null;
            checker = 0;
            excelprocess();
            timer1.Stop();
            for (int b = 0; b <= count; b++)
            { comp[b] = comp1[b]; last[b] = last1[b]; per[b] = per1[b]; }
            timer1.Start();
                i=1;
          /*  if (pic != null)
            {
                timer1.Stop();
            panel1.Visible = false; panel2.Visible = false; panel3.Visible = false;
            panel4.Visible = false; panel5.Visible = false; panel6.Visible = false;
            panel8.Visible = false; panel10.Visible = false;
            panel7.Visible = false; panel9.Visible = false; panel11.Visible = false;
           
            path_img = pic; this.BackgroundImage = new Bitmap(@path_img);
            var pic1 = new Bitmap(this.BackgroundImage, new Size(this.Width, this.Height));
            this.BackgroundImage = pic1;
            comp = null; last = null; per = null;
            checker = 0;
            excelprocess();
           
            for (int b = 0; b <= count; b++)
            { comp[b] = comp1[b]; last[b] = last1[b]; per[b] = per1[b]; }
         Thread.Sleep(150);
         this.BackgroundImage = null;
         panel1.Visible = true; panel2.Visible = true; panel3.Visible = true;
         panel4.Visible = true; panel5.Visible = true; panel6.Visible = true;
         panel8.Visible = true; panel10.Visible = true;
         panel7.Visible = true; panel9.Visible = true; panel11.Visible = true;
         i = 1;
                timer1.Start();
        }
        else
        { comp = null; last = null; per = null;
            checker = 0;
            excelprocess();
            timer1.Stop();
            for (int b = 0; b <= count; b++)
            { comp[b] = comp1[b]; last[b] = last1[b]; per[b] = per1[b]; }
            i = 1;
                timer1.Start(); } */
        }




        
            
   

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            // MessageBox.Show("RunTime ");


     /*       // timer1.Stop();ts.Minutes >= 10//ts.Seconds >= 20
            ts = stopWatch.Elapsed;
            if (ts.Minutes >= 10)
            { //MessageBox.Show("RunTime " + elapsedTime);
                //  MessageBox.Show("RunTime " + ts);
                stopWatch.Stop();
                //timer1.Stop();
                
                comp = null; last = null; per = null;
                checker = 0;
                excelprocess();
              timer1.Stop();
                for (int b = 0; b <= count; b++)
                { comp[b] = comp1[b]; last[b] = last1[b]; per[b] = per1[b]; }
              timer1.Start();
                //MessageBox.Show("RunTime222 ==" + per[5]);
               // timer1.Start();
            }
      

            if (path_img != null && ts.Seconds >= 56 && ts.Seconds <= 60)
            {
                
                if (cnt == 1)
                {
                    panel1.Visible = false; panel2.Visible = false; panel3.Visible = false;
                    panel4.Visible = false; panel5.Visible = false; panel6.Visible = false;
                    panel8.Visible = false; panel10.Visible = false;
                    panel7.Visible = false; panel9.Visible = false; panel11.Visible = false;
                    this.BackgroundImage = new Bitmap(@path_img);
                    var pic = new Bitmap(this.BackgroundImage, new Size(this.Width, this.Height));
                    this.BackgroundImage = pic;
                    cnt = 2;
                }
            }
            else
            {
                if (cnt == 2)
                {this.BackgroundImage = null;
                panel1.Visible = true; panel2.Visible = true; panel3.Visible = true;
                panel4.Visible = true; panel5.Visible = true; panel6.Visible = true;
                panel8.Visible = true; panel10.Visible = true;
                panel7.Visible = true; panel9.Visible = true; panel11.Visible = true;
                cnt = 1;
                } */
            
            if(u == 1)
            panel12.Location = new Point(panel12.Location.X+1, panel12.Location.Y);

            if (panel12.Location.X  >= this.Width)
                u = 0;
            
            if (panel12.Location.X == 1)
            { timer1.Stop(); refresh = 0; refreshfile();  }
                            
                Control[] controls, panel, c;
                Double sign;

            
                // timer1.Start();    
                for (int j = 1; j <= 5; j++)       //select panel
                {

                    panel = this.Controls.Find("panel" + j, true);
                    if (panel.Length == 1)
                    {

                        Panel p = panel[0] as Panel;
                        // if (tmp == 1)
                        //  { temp = p; tmp = 0; cnt++; }
                        if (p != null)
                        {

                            p.Location = new Point(p.Location.X+sp,p.Location.Y);
                            //if (p.Location.Y + p.Height-1 <= 0)
                            if (p.Location.X >= this.Width)
                            {
                                p.Visible = true;
                                if (i > count)
                                { i = 1;
                                //MessageBox.Show(pic);
                                if (pic != null)
                                {
                                   
                                    path_img = pic; panel12.BackgroundImage = new Bitmap(@path_img);
                                    var pic1 = new Bitmap(panel12.BackgroundImage, new Size(panel12.Width, panel12.Height));
                                    panel12.BackgroundImage = pic1;
                                    panel12.Location = new Point(-panel12.Width,panel12.Location.Y); u = 1; panel12.Visible = true;
                                }
                                else
                                    refreshfile();
                                    }
                                for (int k = 1; k <= 5; k++)                 //select label
                                {

                                    controls = this.Controls.Find("Label" + j + "_" + k, true);
                                    if (controls.Length == 1)
                                    {
                                        Label control = controls[0] as Label;
                                        if (control != null)
                                        {
                                           // control.ForeColor = System.Drawing.ColorTranslator.FromHtml(colour); //colour selection
                                            //control.ForeColor = Color(colour);
                                            if (k == 1)
                                            { control.Text = comp[i]; //control.Text = comp[i]; control.ForeColor = System.Drawing.ColorTranslator.FromHtml(colour); 
                                           
                                                control.ForeColor = colour;
                                            }
                                            else if (k == 2)
                                            {
                                                control.Text = last[i];
                                                control.ForeColor = fontcolour;

                                            }
                                            else if (k == 3)
                                            {

                                                if (per[i] < 0.00)
                                                { control.Image = global::WindowsFormsApplication1.Properties.Resources.r; }
                                                else if (per[i] > 0.00)
                                                    control.Image = global::WindowsFormsApplication1.Properties.Resources.g;
                                                else
                                                    control.Image = global::WindowsFormsApplication1.Properties.Resources.y;

                                            }
                                            else if (k == 4)
                                            {
                                                if (per[i] < 0)
                                                    sign = -1;
                                                else
                                                    sign = 1;
                                                control.Text = (per[i] * sign).ToString("0.00") + "%";
                                                control.ForeColor = fontcolour;
                                            }                                             //insert label values
                                            else { //control.BackColor = System.Drawing.ColorTranslator.FromHtml(colour);
                                               
                                                control.BackColor = colour;
                                            }



                                        }
                                    }
                                }
                                i++;
                                

                                p.Location = new Point(-200, 0);
                            }
                        }


                    }
                    Thread.Sleep(speed);
                }

          
        }

        private void label2_4_Click(object sender, EventArgs e)
        {

        }
      
      
       
    }
}
