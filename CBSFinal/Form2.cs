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

namespace CBSFinal
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            mi2 = new MapInfo.MapInfoApplication();
        }
        public static MapInfo.MapInfoApplication mi2;

        public void fill_form()
        {
            label1.Text = MainForm.mi.Eval("iller.il_adi");
            label2.Text = MainForm.mi.Eval("iller.plaka_no");
            label3.Text = "Nüfus Değişim (%): ";
            label4.Text = "Okullaşma Oranı (%): ";
            label5.Text = "Evlilik Değişim (%): ";
            label6.Text = "Suç Değişim (%): ";
            label7.Text = "Suç Puanı: ";
            try
            {
                label3.Text = "Nüfus Değişim (%): " + MainForm.mi.Eval("iller.Nüf_Değişim");
                label4.Text = "Okullaşma Oranı (%): " + MainForm.mi.Eval("iller.Okullaşma_Oranı");
                label5.Text = "Evlilik Değişim (%): " + MainForm.mi.Eval("iller.Evl_Değişim");
                label6.Text = "Suç Değişim (%): " + MainForm.mi.Eval("iller.Suç_Değişim");
                label7.Text = "Suç Puanı: " + MainForm.mi.Eval("iller.Suc_Puan"+MainForm.calcs);
            }
            catch { }

            if (MainForm.calcs>0)
            {
                int p = panel1.Handle.ToInt32();
            //mi2.Do("Open Table Iller as Ilgraph");
            //MainForm.mi.Do("")
            MainForm.mi.Do("Select * from Iller where PLAKA_NO = \"" + MainForm.mi.Eval("iller.plaka_no") + "\" into Selection");
            MainForm.mi.Do("Set Next Document Parent " + panel1.Handle + " Style 1");
            
            
            MainForm.mi.Do("Graph IL_ADI, Suc_Puan"+MainForm.calcs+" ,Evl_Değişim,Okullaşma_Oranı, Suç_Değişim, Nüf_Değişim From Selection Using \"C:\\ProgramData\\MapInfo\\MapInfo\\Professional\\1150\\GraphSupport\\Templates\\Column\\"+MainForm.graphMode+".3tf\" Series In Columns");
            MainForm.mi.Do("Set Graph Title \"Istatistiksel Veriler\" SubTitle \"\" Footnote \"\" TitleGroup \"\" TitleAxisY1 \"Degerler(%)\"");
                label8.Visible = false;//MainForm.graphMode
            }
            else
            {

            }
            // Okullaşma_Oranı, Suç_Değişim, Nüf_Değişim 

            this.ShowDialog();
        }
    }
}
//Graph IL_ADI, Evl_Değişim, Okullaşma_Oranı, Suç_Değişim, Nüf_Değişim From Selection Using "C:\ProgramData\MapInfo\MapInfo\Professional\1150\GraphSupport\Templates\Column\Clustered.3tf" Series In Columns