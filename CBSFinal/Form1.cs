using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MapInfo;
using System.Runtime.InteropServices;
using System.IO;

namespace CBSFinal
{
    public partial class MainForm : Form
    {
        public static MapInfo.MapInfoApplication mi;
        Callback callb;
        public Form2 f2 = new Form2();
        string listpath, win_id, file_path;
        bool n_selected, o_selected, e_selected, s_selected, label_enabled = false, label2_enabled = false, b5pressed=false,graphselected=false;/*, b_selected*/
        int layerLevel=1;
        float evlWeight=1.0f, okulWeight = 1.0f, sucWeight = 1.0f, nufWeight = 1.0f;
        string q;
        public static int calcs=0;
        public static string graphMode= "Clustered";

        [DllImport("user32.dll")]
        static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        public MainForm()
        {
            InitializeComponent();
            
        callb = new Callback(this);
        }
        
        private void nnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("\tEğitimSistem\n \n\t aKpkc, 2017.\n\n Ahmet Kapkiç \n ahmetkapkic@gmail.com", "EğitimSistem Hakkında"); //hakkında butonu
        }
        private void button4_Click(object sender, EventArgs e)
        {
            //nüfus
            //dosya seçtir
            //CurrentStatus.Text = "Initializing";
            openFileDialog1.Title = "Nüfus Verilerinin Olduğu Dosyayı Seçin";     //popupların başlığı
            MessageBox.Show("Veri Konumunu Seçin","Uyarı"); //liste lokasyonu
            DialogResult ListResult = openFileDialog1.ShowDialog();
            if (ListResult == DialogResult.OK)
            {
                removetematik();
                listpath = openFileDialog1.FileName;

                mi.Do("Register Table \"" + listpath + "\"  TYPE XLS Titles Range \"Sheet1!A2:D82\"  Interactive Into \"" + Path.GetDirectoryName(listpath) + "\\Nufus.TAB\"");
                mi.Do("Open Table \"" + Path.GetDirectoryName(listpath) + "\\Nufus.TAB\" Interactive");
                mi.Do("Add Column \"Iller\" (Nüf_Değişim Float)From Nufus Set To Nüf_Değişim Where COL2 = COL1  Dynamic");
                mi.Do("shade window " + win_id + " 1 with Nüf_Değişim ignore 0 ranges apply all use color Brush (2,65280,16777215) -280.2: -18.7 Brush (2,16744576,16777215) Pen (1,2,0) ,-18.7: -9.2 Brush (2,16756816,16777215) Pen (1,2,0) ,-9.2: 0 Brush (2,16764976,16777215) Pen (1,2,0) ,0: 9.8 Brush (2,8453888,16777215) Pen (1,2,0) ,9.8: 1900.4 Brush (2,65280,16777215) Pen (1,2,0) default Brush (2,16777215,16777215) Pen (1,2,0)  # use 0 round 0.1 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 #");
                //shade window 323818504 4 with Değişim_2 ignore 0 ranges apply all use color Brush (2,16744576,16777215)  -280.2: -18.7 Brush (2,16744576,16777215) Pen (1,2,0) ,-18.7: -9.2 Brush (2,16756816,16777215) Pen (1,2,0) ,-9.2: 0 Brush (2,16764976,16777215) Pen (1,2,0) ,0: 9.8 Brush (2,8453888,16777215) Pen (1,2,0) ,9.8: 1900.4 Brush (2,65280,16777215) Pen (1,2,0) default Brush (2,16777215,16777215) Pen (1,2,0)  # use 0 round 0.1 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 # 

                mi.Do("set legend window " + win_id + " layer prev display on shades on symbols off lines off count on title auto Font(\"Arial\", 0, 9, 0) subtitle auto Font(\"Arial\", 0, 8, 0) ascending off ranges Font(\"Arial\", 0, 8, 0) auto display off, auto display on, auto display on, auto display on, auto display on, auto display on");
                n_selected = true;
                layerLevel++;
                button6.Enabled = true;
                checkEnabled();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //okul

            openFileDialog1.Title = "Okul Verilerinin Olduğu Dosyayı Seçin";     //popupların başlığı
            MessageBox.Show("Veri Konumunu Seçin", "Uyarı"); //liste lokasyonu
            DialogResult ListResult = openFileDialog1.ShowDialog();
            if (ListResult == DialogResult.OK)
            {
                removetematik();
                listpath = openFileDialog1.FileName;
                mi.Do("Register Table \"" + listpath + "\"  TYPE XLS Titles Range \"Sheet1!A2:B82\"  Interactive Into \"" + Path.GetDirectoryName(listpath) + "\\Okullasma.TAB\"");
                mi.Do("Open Table \"" + Path.GetDirectoryName(listpath) + "\\Okullasma.TAB\" Interactive");
                mi.Do("Add Column \"Iller\" (Okullaşma_Oranı Float)From Okullasma Set To Okullaşma_Oranı Where COL2 = COL1  Dynamic");
                mi.Do("shade window " + win_id + " 1 with Okullaşma_Oranı ignore 0 ranges apply all use color Brush (2,65280,16777215)  90: 94.7 Brush (2,16744576,16777215) Pen (1,2,0) ,94.7: 96.3 Brush (2,16756816,16777215) Pen (1,2,0) ,96.3: 97.9 Brush (2,16764976,16777215) Pen (1,2,0) ,97.9: 99 Brush (2,8453888,16777215) Pen (1,2,0) ,99: 100 Brush (2,65280,16777215) Pen (1,2,0) default Brush (2,16777215,16777215) Pen (1,2,0)  # use 4 round 0.1 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 #");
                mi.Do("set legend window " + win_id + " layer prev display on shades on symbols off lines off count on title auto Font(\"Arial\", 0, 9, 0) subtitle auto Font(\"Arial\", 0, 8, 0) ascending off ranges Font(\"Arial\", 0, 8, 0) auto display off, auto display on, auto display on, auto display on, auto display on, auto display on");
                //mi.Do("Add Designer Frame Window 237086752 Frame From Layer 1");
                o_selected = true;
                layerLevel++;
                button7.Enabled = true;
                checkEnabled();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //evlilik
            openFileDialog1.Title = "Evlilik/Boşanma Verilerinin Olduğu Dosyayı Seçin";     //popupların başlığı
            MessageBox.Show("Veri Konumunu Seçin", "Uyarı"); //liste lokasyonu
            DialogResult ListResult = openFileDialog1.ShowDialog();
            if (ListResult == DialogResult.OK)
            {
                removetematik();
                listpath = openFileDialog1.FileName;
                mi.Do("Register Table \"" + listpath + "\"  TYPE XLS Titles Range \"Sheet4!A2:H82\"  Interactive Into \"" + Path.GetDirectoryName(listpath) + "\\Evlilik.TAB\"");
                mi.Do("Open Table \"" + Path.GetDirectoryName(listpath) + "\\Evlilik.TAB\" Interactive");
                mi.Do("Add Column \"Iller\" (Evl_Değişim Float)From Evlilik Set To Evl_Değişim Where COL2 = COL1  Dynamic");
                mi.Do("shade window " + win_id + " 1 with Evl_Değişim ignore 0 ranges apply all use color Brush (2,65280,16777215) -0.012: 0.026 Brush (2,65280,16777215) Pen (1,2,0) ,0.026: 0.048 Brush (2, 5308160, 16777215) Pen(1, 2, 0), 0.048: 0.06 Brush(2, 11599616, 16777215) Pen(1, 2, 0), 0.06: 0.071 Brush(2, 16760896, 16777215) Pen(1, 2, 0), 0.071: 0.159 Brush(2, 16744576, 16777215) Pen(1, 2, 0) default Brush(2, 16777215, 16777215) Pen(1, 2, 0)  # use 1 round 0.001 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 #");
                mi.Do("set legend window " + win_id + " layer prev display on shades on symbols off lines off count on title auto Font(\"Arial\", 0, 9, 0) subtitle auto Font(\"Arial\", 0, 8, 0) ascending off ranges Font(\"Arial\", 0, 8, 0) auto display off, auto display on, auto display on, auto display on, auto display on, auto display on");
                e_selected = true;
                layerLevel++;
                button9.Enabled = true;
                checkEnabled();
            }
         }

        private void button3_Click(object sender, EventArgs e)
        {
            
            //suç
            openFileDialog1.Title = "Suç Verilerinin Olduğu Dosyayı Seçin";     //popupların başlığı
            MessageBox.Show("Veri Konumunu Seçin", "Uyarı"); //liste lokasyonu
            DialogResult ListResult = openFileDialog1.ShowDialog();
            if (ListResult == DialogResult.OK)
            {
                removetematik();
                listpath = openFileDialog1.FileName;
                mi.Do("Register Table \"" + listpath + "\"  TYPE XLS Titles Range \"Sheet1!A2:D82\"  Interactive Into \"" + Path.GetDirectoryName(listpath) + "\\Suc.TAB\"");
                mi.Do("Open Table \"" + Path.GetDirectoryName(listpath) + "\\Suc.TAB\" Interactive");
                mi.Do("Add Column \"Iller\" (Suç_Değişim Float)From Suc Set To Suç_Değişim Where COL2 = COL1  Dynamic");
                mi.Do("shade window " + win_id + " 1 with Suç_Değişim ignore 0 ranges apply all use color Brush (2,65280,16777215) -500000: 0 Brush (2,65280,16777215) Pen (1,2,0) ,0: 60 Brush (2,16752800,16777215) Pen (1,2,0) ,60: 200 Brush (2,16744576,16777215) Pen (1,2,0) ,200: 500 Brush (2,16732240,16777215) Pen (1,2,0) ,500: 3490 Brush (2,16719904,16777215) Pen (1,2,0) default Brush (2,16777215,16777215) Pen (1,2,0)  # use 0 round 10 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 #");
                //shade window 323818504 3 with Değişim ignore 0 ranges apply all use all Brush (2,16719904,16777215)  -500000: 0 Brush (2,65280,16777215) Pen (1,2,0) ,0: 60 Brush (2,16752800,16777215) Pen (1,2,0) ,60: 200 Brush (2,16744576,16777215) Pen (1,2,0) ,200: 500 Brush (2,16732240,16777215) Pen (1,2,0) ,500: 3490 Brush (2,16719904,16777215) Pen (1,2,0) default Brush (2,16777215,16777215) Pen (1,2,0)  # use 0 round 10 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 # 

                mi.Do("set legend window " + win_id + " layer prev display on shades on symbols off lines off count on title auto Font(\"Arial\", 0, 9, 0) subtitle auto Font(\"Arial\", 0, 8, 0) ascending off ranges Font(\"Arial\", 0, 8, 0) auto display off, auto display on, auto display on, auto display on, auto display on, auto display on");
                s_selected = true;
                layerLevel++;
                button8.Enabled = true;
                checkEnabled();
            }
        }

        private void helpToolStripButton_Click(object sender, EventArgs e)
        {
            mi.Do("run menu command id 2001");
        }

        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            file_path = ("D:\\Output_" + DateTime.Now.ToFileTime());
            mi.Do("Save Window " + win_id + " As \"" + file_path + ".bmp\" Type \"BMP\" Copyright \"Copyright \" + Chr$(169) + \" 2017, aKpkc.\"");
            MessageBox.Show("Dosya, " + file_path + " konumunda kaydedilmiştir.", "Uyarı"); //liste lokasyonu
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            mi.Do("run menu command 1705");
        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            calculate();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            mi.Do("run menu command 1706");
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
           (e.KeyChar != '.') && (e.KeyChar != '-') && (e.KeyChar != '+') && (e.KeyChar != '/') && (e.KeyChar != '*') && (e.KeyChar != '(') && (e.KeyChar != ')'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            //nd
            textBox1.Text += "ND";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //oo
            textBox1.Text += "OO";
        }

        private void Map_Resize(object sender, EventArgs e)
        {
            if (mi != null)
            {
                // The form has been resized. 
                if (mi.Eval("WindowID(0)") != "")
                {
                    // Update the map to match the current size of the panel. 
                    MoveWindow((System.IntPtr)long.Parse(mi.Eval("WindowInfo(FrontWindow(),12)")), 0, 0, this.Map.Width, this.Map.Height, false);
                }
            }

        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            file_path = ("D:\\Iller_" + DateTime.Now.ToFileTime());
            mi.Do("Export \"Iller\" Into \""+file_path+".csv\" Type \"ASCII\" Delimiter \", \" CharSet \"WindowsTurkish\" Titles");
            MessageBox.Show("Tablo, " + file_path + " konumunda kaydedilmiştir.", "Uyarı"); //liste lokasyonu          
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //ed
            textBox1.Text += "ED";
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //sd
            textBox1.Text += "SD";
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            //open form3


           // q = Legend.Handle.ToString();
            mi.Do("Set Next Document Parent " + q + " Style 3");
            mi.Do("Create Cartographic Legend From Window " + win_id + " Behind Frame From Layer 1"/*+layerLevel*/);
            //mi.Do("set Legend Window " + win_id);
            //mi.Do("Create Legend From Window " +win_id+" Show");
            //lejant
            if (!Legend.Visible)
                Legend.Visible = true;

            else
                Legend.Visible = false;
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            if (!groupBox3.Visible)
            {
                groupBox3.Visible = true;
                groupBox4.Visible = true;
            }
            else
            {
                groupBox3.Visible = false;
                groupBox4.Visible = false;
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            mi.Do("run menu command 1702");
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
           if (!graphselected)
            {
                graphMode = "Custom2";
            }
           else
            {
                graphMode = "Clustered";
            }
            graphselected = !graphselected;
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
        /*    if(!label2_enabled)
            {
                label2_enabled = true;
                //iğrenç bir kod parçası, fix on prod.
            }*/
            if (!label_enabled)
            {
                mi.Do("Set Map Window " + win_id + " Layer "+layerLevel+" Label Auto On");
                label_enabled = true;
            }
            else
            {
                mi.Do("Set Map Window " + win_id + " Layer " + layerLevel + " Label Auto Off");
                label_enabled = false;
            }
            if (b5pressed)
            {
                if (label_enabled)
                {
                    mi.Do("Set Map Window " + win_id + " Layer 2 Label Auto On");
                }
                else
                {
                    mi.Do("Set Map Window " + win_id + " Layer 2 Label Auto Off");
                }
            }
                
        }
        
        private void calculate()
        {
            removetematik();
            /*textbox1 input to asc...
             * oo = (100-Okullaşma_Oranı)*okulWeight
             * ed = Evl_Değişim*evlweight
             * nd = Nüf_Değişim*nufweight
             * sd = Suç_Değişim*sucweight
             * parse info.
             * textbox only accepts +,-,*,/,.,1-9,buttoninputs.
             */
            evlWeight = (float)evlWei.Value;
            okulWeight = (float)okulWei.Value;
            sucWeight = (float)sucWei.Value;
            nufWeight = (float)nufWei.Value;
            string formula = textBox1.Text;
            formula = formula.Replace("OO", "(100-Okullaşma_Oranı)*"+okulWeight)
                .Replace("ED", "Evl_Değişim*"+evlWeight)
                .Replace("ND", "Nüf_Değişim*"+nufWeight)
                .Replace("SD", "Suç_Değişim*"+sucWeight).Replace(',','.');
            calcs++;
            //hesapla, veriler tam değilse disabled
            //            mi.Do("Add Column \"Iller\" (DYPp Float)From _91_Sec_Il_Mer Set To DYP*100/GECERLI_OY Where COL2 = COL1  Dynamic");
            //mi.Do("Alter Table \"Iller\"(add Suc_Puan Float) Interactive");
            //mi.Do("Add Column \"Iller\"(Suc_Puan Float) From Iller Set To ((100 - Okullaşma_Oranı)*okulWeight) * (Evl_Değişim*" + evlWeight + ") * (Suç_Değişim*" + sucWeight + " /( Nüf_Değişim*" + nufWeight + "))");
            //if (!b5pressed)
            if (formula=="")
            {
                formula= "OO * ED * SD / ND";
                MessageBox.Show("Formül kısmı boş bırakıldığından varsayılan formül olan (OO*ED*SD/ND) kullanılmıştır.", "Uyarı");
            }
            mi.Do("Add Column \"Iller\"(Suc_Puan"+calcs+" Float) From Iller Set To "+formula);
            //mi.Do("Update Iller Set Suc_Puan = (100 - Okullaşma_Oranı)) * (Evl_Değişim + 1) * (Suç_Değişim / Nüf_Değişim)");
            mi.Do("shade window " + win_id + " 1 with Suc_Puan" + calcs + " ignore 0 ranges apply all use color Brush (2,65280,16777215) -500000: 0 Brush (2,65280,16777215) Pen (1,2,0) ,0: 60 Brush (2,16752800,16777215) Pen (1,2,0) ,60: 200 Brush (2,16744576,16777215) Pen (1,2,0) ,200: 500 Brush (2,16732240,16777215) Pen (1,2,0) ,500: 93490 Brush (2,16719904,16777215) Pen (1,2,0) default Brush (2,16777215,16777215) Pen (1,2,0)  # use 0 round 10 inflect on Brush (2,16776960,16777215) at 3 by 0 color 1 #");
            b5pressed = true;
            mi.Do("Add Map Window " + win_id + " Auto Layer Iller");
            mi.Do("Set Map Window " + win_id + " Order Layers 1 DestGroupLayer 0 Position 2");
            mi.Do("Set Map Window " + win_id + " Layer 2 Label With Suc_Puan" + calcs + " Position Below");
            mi.Do("Set Map Window " + win_id + " Layer 2 Selectable Off");
            layerLevel += 2;
            if (label_enabled)
            {
                mi.Do("Set Map Window " + win_id + " Layer 2 Label Auto On");
                
            }
            if(layerLevel>3)
            {
                mi.Do("Set Map Window "+win_id+" Layer 3 Label Auto Off");
            }
            
            
            
            //Add Map Window 232396288 Auto Layer Iller
            //Set Map Window 232396288 Order Layers 1 DestGroupLayer 0 Position 2
            //Set Map Window 232396288  Layer 2 Label With SUÇ Position Below
           // toolStripButton8.Enabled = false;
            
        }
        private void button6_Click(object sender, EventArgs e)
        {
            //bilgi
            mi.Do("run menu command id 2001");
        }
        private void checkEnabled()
        {
            if (n_selected && o_selected && e_selected && s_selected /*&& b_selected*/)
            {
                toolStripButton8.Enabled = true;

            }
            else { toolStripButton8.Enabled = false; }
            
        }
        public void removetematik()
        {
            for (int k = Convert.ToInt16(mi.Eval("mapperinfo(" + win_id + ",9)")); k > 0; k = k - 1)
            {
                if (Convert.ToInt16(mi.Eval("layerinfo(" + win_id + "," + Convert.ToString(k) + ",24)")) >= 3) //katman 3 ise kaldır
                {
                    mi.Do("remove map layer \"" + mi.Eval("layerinfo(" + win_id + "," + Convert.ToString(k) + ",1)") + "\"");
                    layerLevel--;
                }
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            mi = new MapInfo.MapInfoApplication();
            int p = Map.Handle.ToInt32();
            mi.Do("set next document parent " + p.ToString() + "style 1");
            mi.Do("set application window " + p.ToString());
            mi.Do("run application \"" + "d:/Final.wor" + "\"");
            win_id = mi.Eval("frontwindow()");
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory); //temel başlangıç konumları
            openFileDialog1.Filter = "Excel files (*.xls)|*.xls"; //[1]

            mi.SetCallback(callb);
            mi.Do("create buttonpad \"a\" as toolbutton calling OLE \"info\" id 2001");
            mi.Do("Set Map Window " + win_id + " Layer 1 Label With IL_ADI Position Above");
            q = p.ToString();
        }
        //textbox1 bina sayısı removed
        //Map panel
        //MainForm form
        //eklenen layer tekrar eklenince üzerine yazma çabaları oluyor. olmasın.
        //aralık göstermiyoruz, olsun. seviyeleri de belli değil.
    }
    
    [ComVisible(true)]
    public class Callback
    {
        MainForm f1;

        public Callback(MainForm _f1)
        {
            f1 = _f1;
        }
        public void info(string a)
        {
            try{ 

            int k = Convert.ToInt32(MainForm.mi.Eval("searchpoint(frontwindow(),commandinfo(1),commandinfo(2))"));
            string tabloadi = "";
            for (int i = 1; i <= k; i++)
            {
                tabloadi = MainForm.mi.Eval("SearchInfo(" + i.ToString() + ",1)");
                String row_id = MainForm.mi.Eval("SearchInfo(" + i.ToString() + ",2)");
                MainForm.mi.Do("Fetch rec " + row_id + " From " + tabloadi);
                if ((tabloadi == "Iller"))
                {
                    f1.Invoke(new mapinfo(f1.f2.fill_form));
                }
            }
        }
            catch { }
        }
        delegate void mapinfo();
    }
}