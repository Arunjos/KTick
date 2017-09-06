using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Security.Cryptography;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public String pic=null;
        int flag = 0,zen;
        String wrt="";
        public Form2 form2 = new Form2();
        static byte[] bytes = ASCIIEncoding.ASCII.GetBytes("arunjose");
        public Form1()
        {
            InitializeComponent();
            //timer1.Start();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

           /*StreamReader sr1 = new StreamReader(@".\file1.txt");
             string  line1 = sr1.ReadLine();
             sr1.Close();
            try
            {
             DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
             MemoryStream memoryStream = new MemoryStream(Convert.FromBase64String(line1));
             CryptoStream cryptoStream = new CryptoStream(memoryStream, cryptoProvider.CreateDecryptor(bytes, bytes), CryptoStreamMode.Read);
             StreamReader reader = new StreamReader(cryptoStream);
               string tyms = reader.ReadToEnd();
                zen = int.Parse(tyms);
               }
               catch (Exception xex)
               {
                   zen = 30;
                   MessageBox.Show("Trial expierd");
                   Close();
               }

           
           if (zen >= 13)
           {
               MessageBox.Show("12 Trial Run Over");
               Close();
               //Application.Exit();

           }
           else
               if (zen < 13)
           {
               MessageBox.Show("It is a Trial Version RUNS REMAINING: "+(13 - zen));
               zen = zen + 1;
              
               zen.ToString();
               DESCryptoServiceProvider cryptoProvider1 = new DESCryptoServiceProvider();
               MemoryStream memoryStream1 = new MemoryStream();
               CryptoStream cryptoStream1 = new CryptoStream(memoryStream1, cryptoProvider1.CreateEncryptor(bytes, bytes), CryptoStreamMode.Write);
               StreamWriter writer = new StreamWriter(cryptoStream1);
               writer.Write(zen);
               writer.Flush();
               cryptoStream1.FlushFinalBlock();
               writer.Flush();
               string store = Convert.ToBase64String(memoryStream1.GetBuffer(), 0, (int)memoryStream1.Length);

               System.IO.File.WriteAllText(@".\file1.txt", string.Empty);
               StreamWriter sw1 = new StreamWriter(@".\file1.txt", true, Encoding.ASCII);
               sw1.Write(store);
               sw1.Close();
          */
            try
            {

                StreamReader sr = new StreamReader(@".\file2.txt");


                pic = sr.ReadLine();


                if (pic != null)
                {  //MessageBox.Show("STRING IS" +line );
                    if (File.Exists(pic) != true)
                    {
                        pic = null;
                    }
                }


                //close the file
                sr.Close();

            }
            catch (Exception xex)
            {
                MessageBox.Show("Exception: " + xex.Message);
            } //

               String line;
               try
               {

                   StreamReader sr = new StreamReader(@".\file.txt");


                   line = sr.ReadLine();


                   if (line != null)
                   {  //MessageBox.Show("STRING IS" +line );
                       if (File.Exists(line) == true)
                       {
                           this.button1.Text = "";
                           this.button1.BackgroundImage = global::WindowsFormsApplication1.Properties.Resources.downloadpp;
                          // form2.Tag = this;
                           form2.Show();
                           this.Show();
                           form2.start(line, pic);
                       }
                       else
                       {
                          // form2.Tag = this;
                           form2.Show();
                           this.Show();
                       }
                   }
                   else
                   {
                      // form2.Tag = this;
                       form2.Show();
                       this.Show();
                   }

                   //close the file
                   sr.Close();

               }
               catch (Exception xex)
               {
                   MessageBox.Show("Exception: " + xex.Message);
               }

              // form2.Tag = this;
              // form2.Show(this);
               // Hide();
          //}
        
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
           wrt = form2.excelfile();
           System.IO.File.WriteAllText(@".\file.txt", string.Empty);
           try
           {
               
               StreamWriter sw = new StreamWriter(@".\file.txt", true, Encoding.ASCII);

               sw.Write(wrt);
         

               //close the file
               sw.Close();
           }
           catch (Exception ex)
           {
               Console.WriteLine("Exception: " + ex.Message);
           }
           this.button1.Text = " ";
           this.button1.BackgroundImage = global::WindowsFormsApplication1.Properties.Resources.play;
           flag = 0;

        }

  
        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            form2.speeds(Convert.ToInt32(trackBar1.Value));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
               if ( this.button1.Text == " ")
                {
                   //
                    //MessageBox.Show("PLEASE SELECT ANddddY FILE");
                        
                        if (flag == 0)
                        {
                            if (wrt == "")
                                MessageBox.Show("PLEASE SELECT ANY FILE");
                            else
                            {
                                this.button1.Text = "";
                                this.button1.BackgroundImage = global::WindowsFormsApplication1.Properties.Resources.downloadpp;
                                form2.start(wrt, pic);
                               // form2.excelprocess();
                            }
                        }
                        else
                        {
                            form2.starttimer(); flag = 0;
                            this.button1.Text = "";
                            this.button1.BackgroundImage = global::WindowsFormsApplication1.Properties.Resources.downloadpp;
                        }
                    
                  

                }
                else
               {
                   this.button1.Text = " ";
                    this.button1.BackgroundImage = global::WindowsFormsApplication1.Properties.Resources.play;
                    flag = 1;
                    form2.tymstopper();

                }
            
        }

     

      /*  private void button6_Click(object sender, EventArgs e)
        {
            form2.image();
        }*/

        /*private void button7_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            { MessageBox.Show("PLEASE SELECT ANY FILE"); }
            else
            {
                if (this.button1.Text == "PAUSE")
                {
                    this.button1.BackColor = System.Drawing.Color.DarkOliveGreen;
                    this.button1.Text = "PLAY";
                    flag = 1;
                    form2.tymstopper();
                }
                form2.flash();
            }
        }*/

        private void button3_Click(object sender, EventArgs e)
        {
            String ab = form2.img_file();
            System.IO.File.WriteAllText(@".\file2.txt", string.Empty);
            try
            {

                StreamWriter sw = new StreamWriter(@".\file2.txt", true, Encoding.ASCII);

                sw.Write(ab);


                //close the file
                sw.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            //textBox3.Text = form2.flash_file();
            
                DialogResult res = colorDialog1.ShowDialog();
                if (res.Equals(DialogResult.OK))
                {
                    form2.Color(colorDialog1.Color);
                  //  textBox3.BackColor = colorDialog1.Color;
                   
                   //textBox3.Text =Convert.ToString(colorDialog1.Color);
                }
            }
        
        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //form2.Color(comboBox1.Text);
        }

        private void openExcelFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button2_Click(null, null);
        }

        private void openImageFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button3_Click(null, null);
        }

        private void changeFontColourToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button4_Click(null,null);
        }

        private void removeImageToolStripMenuItem_Click(object sender, EventArgs e)
        {
           form2.rmpic();
        }

        private void lowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            form2.speeds(3);
        }

        private void mediumToolStripMenuItem_Click(object sender, EventArgs e)
        {
            form2.speeds(2);
        }

        private void lowToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            form2.speeds(1);
        }

        private void startPauseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button1_Click(null,null);
        }

        private void helpToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Update Soon...");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Available Soon...");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult res = colorDialog2.ShowDialog();
            if (res.Equals(DialogResult.OK))
            {
                form2.fontColor(colorDialog2.Color);
                //  textBox3.BackColor = colorDialog1.Color;

                //textBox3.Text =Convert.ToString(colorDialog1.Color);
            }
        }

        private void valueFontColourToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button6_Click(null, null);
        }

     

  


    }
}
