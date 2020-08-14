using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Net;
using System.Net.Mail;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Collections;

namespace PMYS
{
    public partial class Personel : Form
    {
        public Personel()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
               // timer1.Start();
                baglan();
                PersonelGetir();
                DonemBilgileriGetir();
                HareketBilgileriniGetir();
                mtHareketGiris.Text = "09.00";
                mtHareketCıkıs.Text = "18.00";
                comboboxvericek();
                GunleriGetir();
            }
            catch (Exception q )
            {
                MessageBox.Show(q.Message);
            }
        }
        public SqlConnection baglanti;
        private void baglan()
        {
            try
            {
                baglanti = new SqlConnection("Data Source=CASPER\\SQLEXPRESS; Initial Catalog=PMYSis; Integrated Security=true");
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
            }
            catch (Exception q)
            {   
               MessageBox.Show(q.Message);
            }
        }
        private void PersonelGetir()
        {
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                SqlCommand cmd = new SqlCommand("SELECT Personel_id,Tc_kimlik_no,adi,soyadi,E_Posta FROM Personel", baglanti);
                SqlDataAdapter adp = new SqlDataAdapter(cmd);
                System.Data.DataTable dtable = new System.Data.DataTable();
                adp.Fill(dtable);
                dgwPersListe.DataSource = dtable;
                baglanti.Close();
            }
            catch (Exception q)
            {
                 MessageBox.Show(q.Message);
            }
        }
        private void GunleriGetir()
        {
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                SqlCommand cmd = new SqlCommand("SELECT * from Donem_Gun", baglanti);
                SqlDataAdapter adp = new SqlDataAdapter(cmd);
                System.Data.DataTable dtable = new System.Data.DataTable();
                adp.Fill(dtable);
                dgwGunListele.DataSource = dtable;
                baglanti.Close();
            }
            catch (Exception q)
            {   
               MessageBox.Show(q.Message);
            }
        }
        private void DonemBilgileriGetir()
        {
            try
            {
                if (baglanti.State == ConnectionState.Closed) // Veri tabanına bağlantı kapalı ise , açıyor.
                {
                    baglanti.Open();
                }
                SqlCommand cmd = new SqlCommand("SELECT Donem_id,DonemAdi,DonemYili,DonemAyi,Durumu FROM Donem", baglanti);
                SqlDataAdapter adp = new SqlDataAdapter(cmd);
                System.Data.DataTable dtable = new System.Data.DataTable();
                adp.Fill(dtable);
                dgwDnmListe.DataSource = dtable;
                baglanti.Close();
            }
            catch (Exception q)
            {   
            MessageBox.Show(q.Message);
            }
        }
        private void HareketBilgileriniGetir()
        {
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                SqlCommand cmd = new SqlCommand("Select Hareket.Hareket_id,Personel.adi,Personel.soyadi,Donem_Gun.Gun,Hareket.GirisSaati,Hareket.CıkısSaati from Personel,Donem_Gun,Hareket where (Personel.Personel_id=Hareket.Personel_id and Donem_Gun.Gun_id=Hareket.Gun_id)", baglanti);
                SqlDataAdapter adp = new SqlDataAdapter(cmd);
                System.Data.DataTable dtable = new System.Data.DataTable();
                adp.Fill(dtable);
                dgwHareketListe.DataSource = dtable;
                baglanti.Close();
            }
            catch (Exception q)
            {       
            MessageBox.Show(q.Message);
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("INSERT INTO Personel (Tc_kimlik_no,adi,soyadi,E_Posta) VALUES (@kmlk,@ad,@soyad,@eposta)", baglanti);
                cmd.Parameters.AddWithValue("@ad", txtPersAd.Text);
                cmd.Parameters.AddWithValue("@soyad", txtPersSoyad.Text);
                cmd.Parameters.AddWithValue("@eposta", txtPersEPosta.Text);
                cmd.Parameters.AddWithValue("@kmlk", txtTcno.Text);
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                cmd.ExecuteNonQuery();
                baglanti.Close();
                PersonelGetir();
                txtPersAd.Clear();
                txtPersSoyad.Clear();
                txtPersEPosta.Clear();
                txtTcno.Clear();
                label17.Text = "";
            }
            catch (Exception q)
            {   
               MessageBox.Show(q.Message);
            }  
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("UPDATE Personel SET Tc_kimlik_no=@kmlk,adi=@ad,soyadi=@soyad,E_Posta=@eposta WHERE Personel_id=@id ", baglanti);
              cmd.Parameters.AddWithValue("@kmlk",txtTcno.Text); 
                cmd.Parameters.AddWithValue("@id", dgwPersListe.CurrentRow.Cells[0].Value);
                cmd.Parameters.AddWithValue("@ad", txtPersAd.Text);
                cmd.Parameters.AddWithValue("@soyad", txtPersSoyad.Text);
                cmd.Parameters.AddWithValue("@eposta", txtPersEPosta.Text);
              
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                cmd.ExecuteNonQuery();
                baglanti.Close();
            PersonelGetir();
            }
            catch (Exception q)
            {       
                MessageBox.Show(q.Message);
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
         
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("DELETE FROM Personel WHERE Personel_id=@id", baglanti);
                cmd.Parameters.AddWithValue("@id", dgwPersListe.CurrentRow.Cells[0].Value);
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                cmd.ExecuteNonQuery();
                baglanti.Close();
                PersonelGetir();
            }
            catch (Exception q)
            {       
                 MessageBox.Show(q.Message);
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            string df = "";
            int t = 0;
            string c = "";
            try
            {
                #region Dönemekle
		
                if (cmbDonemDurum.Text != "")
                {

                    SqlCommand cmd = new SqlCommand("INSERT INTO Donem (DonemAdi,DonemYili,DonemAyi,Durumu) VALUES (@dnmad,@dnmyil,@dnmay,@Durum)", baglanti);
                    cmd.Parameters.AddWithValue("@dnmad", txtDnmAdi.Text);
                    cmd.Parameters.AddWithValue("@dnmyil", cmbDnmYili.Text);
                    cmd.Parameters.AddWithValue("@dnmay", cmbDnmAyi.Text);
                    cmd.Parameters.AddWithValue("@Durum",cmbDonemDurum.Text);
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    cmd.ExecuteNonQuery();
                    baglanti.Close();
                    DonemBilgileriGetir();
                }
                else MessageBox.Show("Alanlar Boş Olamaz.");
                btnDnmGunclle.Visible = true;
                btnDnmSil.Visible = true;
                btnDnmKaydet.Visible = false; 
	#endregion

                comboboxvericek();
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                SqlCommand cmdd = new SqlCommand("Select d.DonemYili,d.DonemAyi,d.Donem_id from Donem d where d.Donem_id=@dnmid ", baglanti);
                cmdd.Parameters.AddWithValue("@dnmid", cmbGunDnmAdi.SelectedValue.ToString()); //dönemidkısmı düzenlenecek
                SqlDataReader dr = cmdd.ExecuteReader();

                cmbHareketGun.Items.Clear();
                while (dr.Read())
                {
                    df = dr["DonemYili"].ToString();
                    c = dr["DonemAyi"].ToString();

                }
                dr.Close();
                //       MessageBox.Show(df);
                if (c == "Ocak") t = 01;
                if (c == "Şubat") t = 02;
                if (c == "Mart") t = 03;
                if (c == "Nisan") t = 04;
                if (c == "Mayıs") t = 05;
                if (c == "Haziran") t = 06;
                if (c == "Temmuz") t = 07;
                if (c == "Ağustos") t = 08;
                if (c == "Eylül") t = 09;
                if (c == "Ekim") t = 10;
                if (c == "Kasım") t = 11;
                if (c == "Aralık") t = 12;
                //     MessageBox.Show(t.ToString());

                int gun;
                gun = DateTime.DaysInMonth(Convert.ToInt32(df), Convert.ToInt32(t));
                //   MessageBox.Show(gun.ToString());
                cmbGun.Items.Clear();
                
                for (int yaz = 0; yaz <= gun-1; yaz++)
                {
                    SqlCommand gunkayt = new SqlCommand("insert into Donem_Gun (Gun,Donem_id) values (@gun,@Dnm_id)", baglanti);
                    gunkayt.Parameters.AddWithValue("@gun", yaz+1);
                    gunkayt.Parameters.AddWithValue("@Dnm_id", cmbGunDnmAdi.SelectedValue.ToString());
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    gunkayt.ExecuteNonQuery();
                    baglanti.Close();
                }
                GunleriGetir();
            }
            catch (Exception q)
            {
                MessageBox.Show(q.Message);
            }  
        }
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                if (cmbDonemDurum.Text != "")
                {
                    SqlCommand cmd = new SqlCommand("UPDATE Donem SET DonemAdi=@dnmad,DonemYili=@dnmyil,DonemAyi=@dnmay,Durumu=@Durum WHERE Donem_id=@id ", baglanti);
                    cmd.Parameters.AddWithValue("@id", dgwDnmListe.CurrentRow.Cells[0].Value);
                    cmd.Parameters.AddWithValue("@dnmad", txtDnmAdi.Text);
                    cmd.Parameters.AddWithValue("@dnmyil", cmbDnmYili.Text);
                    cmd.Parameters.AddWithValue("@dnmay", cmbDnmAyi.Text);
                    cmd.Parameters.AddWithValue("@Durum",cmbDonemDurum.Text);
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    cmd.ExecuteNonQuery();
                    baglanti.Close();
                    DonemBilgileriGetir();
                }
                else MessageBox.Show("Alanlar Boş Olamaz.");
            } 
            catch (Exception q)
            {       
                 MessageBox.Show(q.Message);
            }
        }
        private void button4_Click_1(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("DELETE FROM Donem WHERE Donem_id=@id", baglanti);
                cmd.Parameters.AddWithValue("@id", dgwDnmListe.CurrentRow.Cells[0].Value);
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                cmd.ExecuteNonQuery();
                baglanti.Close();
                DonemBilgileriGetir();
            }
            catch (Exception q)
            {
                 MessageBox.Show(q.Message);
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("UPDATE Hareket SET GirisSaati=@grs,CıkısSaati=@cks,Personel_id=@pers,Gun_id=@dnm WHERE Personel_id=@id ", baglanti);
                cmd.Parameters.AddWithValue("@id", cmbHareketAdiSoyadi.SelectedValue.ToString());
                cmd.Parameters.AddWithValue("@grs", mtHareketGiris.Text);
                cmd.Parameters.AddWithValue("@cks", mtHareketCıkıs.Text);

                cmd.Parameters.AddWithValue("@pers", cmbHareketAdiSoyadi.SelectedValue.ToString());
                cmd.Parameters.AddWithValue("@dnm", cmbHareketGun.SelectedValue.ToString());
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                cmd.ExecuteNonQuery();
                baglanti.Close();
                HareketBilgileriniGetir();
            }
            catch (Exception q)
            {       
                 MessageBox.Show(q.Message);
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("DELETE FROM Hareket WHERE Hareket_id=@id", baglanti);
                cmd.Parameters.AddWithValue("@id", dgwHareketListe.CurrentRow.Cells[0].Value);
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                cmd.ExecuteNonQuery();
                baglanti.Close();
                HareketBilgileriniGetir();
            }
            catch (Exception q)
            {       
                MessageBox.Show(q.Message);
            }
        }
        private void personelveriyukle()
        {
            try
            {
                dgwPersListe.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                txtPersAd.Text = dgwPersListe.CurrentRow.Cells[2].Value.ToString();
                txtPersSoyad.Text = dgwPersListe.CurrentRow.Cells[3].Value.ToString();
                txtPersEPosta.Text = dgwPersListe.CurrentRow.Cells[4].Value.ToString();
                txtTcno.Text=dgwPersListe.CurrentRow.Cells[1].Value.ToString();
            }
            catch (Exception q )
            {   
                 MessageBox.Show(q.Message);
            }
        }
        private void donembilgisiyukle()
        {
            try
            {
                dgwDnmListe.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                txtDnmAdi.Text = dgwDnmListe.CurrentRow.Cells[1].Value.ToString();
                cmbDnmYili.Text = dgwDnmListe.CurrentRow.Cells[2].Value.ToString();
                cmbDnmAyi.Text = dgwDnmListe.CurrentRow.Cells[3].Value.ToString();
                cmbDonemDurum.Text = dgwDnmListe.CurrentRow.Cells[4].Value.ToString();
            }
            catch (Exception q)
            {   
                MessageBox.Show(q.Message);
            }
        }
        private void hareketbilgileriniyukle()
        {
            try
            {
                dgwHareketListe.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                mtHareketGiris.Text = dgwHareketListe.CurrentRow.Cells[4].Value.ToString();
                mtHareketCıkıs.Text = dgwHareketListe.CurrentRow.Cells[5].Value.ToString();
             //   cmbHareketAdiSoyadi.Text = dgwHareketListe.CurrentRow.Cells[1].Value.ToString() ;
            //    cmbHareketDnmAdi.Text = dgwHareketListe.CurrentRow.Cells[2].Value.ToString();

            }
            catch (Exception q)
            {       
                MessageBox.Show(q.Message);
            }
        }
        private void gunbilgisiyukle() 
        {
            try
            {
                dgwGunListele.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                cmbGun.Text = dgwGunListele.CurrentRow.Cells[1].Value.ToString();
            }
            catch (Exception q)
            {       
                MessageBox.Show(q.Message);
            }
        }
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                donembilgisiyukle();
                btnDnmKaydet.Visible = false;
                btnDnmKaydet.Enabled = true;
                btnDnmSil.Visible = true;
                btnDnmSil.Enabled = true;
                btnDnmGunclle.Visible = true;
                btnDnmGunclle.Enabled = true;
                txtDnmAdi.Enabled = true;
                cmbDnmYili.Enabled = true;
                cmbDnmAyi.Enabled = true;
                cmbDonemDurum.Enabled = true;
            }
            catch (Exception q)
            {   
                 MessageBox.Show(q.Message);
            }
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                personelveriyukle();
                txtPersAd.Enabled = true;
                txtPersSoyad.Enabled = true;
                txtPersEPosta.Enabled = true;
                btnPersGunclle.Enabled = true;
                btnPersSil.Enabled = true;
                txtTcno.Enabled = true;
                btnPersKaydet.Enabled = false;
            }
            catch (Exception q)
            {   
                 MessageBox.Show(q.Message);
            }
        }
        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
               // hareketbilgileriniyukle();
                btnHareketKaydet.Enabled = false;
                btnHareketGuncelle.Enabled = true;
                btnHareketSil.Enabled = true;
                mtHareketGiris.Enabled = true;
                mtHareketCıkıs.Enabled = true;
                cmbHareketAdiSoyadi.Enabled = true;
                cmbHareketDnmAdi.Enabled = true;
                cmbHareketGun.Enabled = true;
            }
            catch (Exception q)
            {       
                 MessageBox.Show(q.Message);
            }
           
        }
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                txtPersAd.Enabled = true;
                txtPersSoyad.Enabled = true;
                txtPersEPosta.Enabled = true;
                btnPersKaydet.Enabled = true;
                txtTcno.Enabled = true;
                btnPersGunclle.Enabled = false;
                btnPersSil.Enabled = false;
                /*************************************************************************************************************/
                txtPersAd.Clear();
                txtPersSoyad.Clear();
                txtPersEPosta.Clear();
                txtTcno.Clear();
                label17.Text = "";
            }
            catch (Exception q)
            {       
                 MessageBox.Show(q.Message);
            }
        }
        private void comboboxvericek()
        {
            try
            {
                System.Data.DataTable table = new System.Data.DataTable();
                System.Data.DataTable table2 = new System.Data.DataTable();
                System.Data.DataTable gunler = new System.Data.DataTable();
                SqlDataAdapter da = new SqlDataAdapter("select adi+ ' '+soyadi as ad , Personel_id from Personel", baglanti);
                SqlDataAdapter da2 = new SqlDataAdapter("select * from Donem where Durumu='Aktif' order by Donem_id desc", baglanti);
               // SqlDataAdapter gunlerdatatable = new SqlDataAdapter("Select g.Gun from Donem_Gun g where g.Donem_id=9", baglanti);

                //SqlDataAdapter gunlerdatatable = new SqlDataAdapter("Select * from Donem_Gun ", baglanti);
                //gunlerdatatable.Fill(gunler);
                da.Fill(table);
                da2.Fill(table2);
                cmbHareketAdiSoyadi.DataSource = new BindingSource(table, null);
                cmbHareketAdiSoyadi.DisplayMember = "ad";
                cmbHareketAdiSoyadi.ValueMember = "Personel_id";
                /*************************************************************************************************************/
                cmbHareketDnmAdi.DataSource = new BindingSource(table2, null);
                cmbHareketDnmAdi.DisplayMember = "DonemAdi";
                cmbHareketDnmAdi.ValueMember = "Donem_id";
                /*************************************************************************************************************/
                cmbGunDnmAdi.DataSource = new BindingSource(table2, null);
                cmbGunDnmAdi.DisplayMember = "DonemAdi";
                cmbGunDnmAdi.ValueMember = "Donem_id";
                /*************************************************************************************************************/
                cmbRaporlamaDonemSec.DataSource = new BindingSource(table2, null);
                cmbRaporlamaDonemSec.DisplayMember = "DonemAdi";
                cmbRaporlamaDonemSec.ValueMember = "Donem_id";
                /*************************************************************************************************************/
                cmbRaporlamaPersSec.DataSource = new BindingSource(table, null);
                cmbRaporlamaPersSec.DisplayMember = "ad";
                cmbRaporlamaPersSec.ValueMember = "Personel_id";
                /*************************************************************************************************************/
                //cmbHareketGun.DataSource = new BindingSource(gunler, null);
                //cmbHareketGun.DisplayMember = "Gun";
                //cmbHareketGun.ValueMember = "Gun_id";
            }
            catch (Exception q)
            {    
                MessageBox.Show(q.Message);
            }
        }
        private void cmbHareketAdiSoyadi_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        int sayi = 0;
        private void cmbHareketDnmAdi_SelectedIndexChanged(object sender, EventArgs e)
        {

            sayi++;
            if (sayi > 2)
            {
                cmbHareketGun.Text = "";
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                SqlCommand cmdd = new SqlCommand("Select g.Gun from Donem_Gun g where g.Donem_id=@dnmid ", baglanti);
                cmdd.Parameters.AddWithValue("@dnmid", cmbHareketDnmAdi.SelectedValue.ToString());
                SqlDataReader dr = cmdd.ExecuteReader();
                String a;
                cmbHareketGun.Items.Clear();
                while (dr.Read())
                {
                    a = dr["Gun"].ToString();
                    cmbHareketGun.Items.Add(a);
                }
                dr.Close();
            }
            
        }
        private void btnYenile_Click(object sender, EventArgs e)
        {
            comboboxvericek();
        }
        private void cmbDnmGun_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
        private void btnGunYenile_Click(object sender, EventArgs e)
        {
            comboboxvericek();
        }
        private void btnGun_Ekle_Click_1(object sender, EventArgs e)
        {
            try
            {
                SqlCommand gunkayt = new SqlCommand("insert into Donem_Gun (Gun,Donem_id) values (@gun,@Dnm_id)", baglanti);
                gunkayt.Parameters.AddWithValue("@gun", cmbGun.Text);
                gunkayt.Parameters.AddWithValue("@Dnm_id", cmbGunDnmAdi.SelectedValue.ToString());
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                gunkayt.ExecuteNonQuery();
                baglanti.Close();
                
                GunleriGetir();
             }
            catch (Exception q)
            {
                MessageBox.Show(q.Message);
            }
        }
        private void btnGunYenile_Click_1(object sender, EventArgs e)
        {
            comboboxvericek();
            GunleriGetir();
        }
        private void btnGun_Guncelle_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("UPDATE Donem_Gun SET Gun=@gun,Donem_id=@gid WHERE Gun_id=@id ", baglanti);
                cmd.Parameters.AddWithValue("@id", dgwGunListele.CurrentRow.Cells[0].Value);
                cmd.Parameters.AddWithValue("@gun", cmbGun.Text);
                cmd.Parameters.AddWithValue("@gid", cmbGunDnmAdi.SelectedValue.ToString());
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                cmd.ExecuteNonQuery();
                baglanti.Close();
                GunleriGetir(); ;
           
                }
            catch (Exception q)
            {       
                 MessageBox.Show(q.Message);
            }
        }
        private void btnGunSil_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = new SqlCommand("DELETE FROM Donem_Gun WHERE Gun_id=@id", baglanti);
                cmd.Parameters.AddWithValue("@id", dgwGunListele.CurrentRow.Cells[0].Value);
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                cmd.ExecuteNonQuery();
                baglanti.Close();
                GunleriGetir();
                }
            catch (Exception q)
            {       
                 MessageBox.Show(q.Message);
            }
        }
        private void btnHareketKaydet_Click(object sender, EventArgs e)
        {
            try
            {
                 string a="";
                #region MyRegion
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                SqlCommand cmdd = new SqlCommand("Select g.Gun_id from Donem d,Donem_Gun g where g.Gun='"+cmbHareketGun.Text+"' and d.Donem_id=g.Donem_id ", baglanti);
                cmdd.Parameters.AddWithValue("@dnmid", cmbHareketDnmAdi.SelectedValue.ToString());
                SqlDataReader dr = cmdd.ExecuteReader();
               
                while (dr.Read())
                {
                    a = dr["Gun_id"].ToString();
                    
                }
                dr.Close();
                MessageBox.Show(a);

                #endregion



                SqlCommand cmd = new SqlCommand("INSERT INTO Hareket (GirisSaati,CıkısSaati,Personel_id,Gun_id) VALUES (@grs,@cks,@persid,@gunid)", baglanti);
                cmd.Parameters.AddWithValue("@grs", mtHareketGiris.Text);
                cmd.Parameters.AddWithValue("@cks", mtHareketCıkıs.Text);
                cmd.Parameters.AddWithValue("@persid", cmbHareketAdiSoyadi.SelectedValue.ToString());
                cmd.Parameters.AddWithValue("@gunid", a);
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                cmd.ExecuteNonQuery();
                baglanti.Close();
                HareketBilgileriniGetir();
                 }
            catch (Exception q)
            {       
                 MessageBox.Show(q.Message);
            }
        }
        private void btnDnmYeni_Click(object sender, EventArgs e)
        {
            try
            {
                btnDnmKaydet.Enabled = true;
                btnDnmKaydet.Visible = true;
                btnDnmGunclle.Visible = false;
                btnDnmSil.Visible = false;
                txtDnmAdi.Enabled = true;
                cmbDnmYili.Enabled = true;
                cmbDnmAyi.Enabled = true;
                cmbDonemDurum.Enabled = true;
               txtDnmAdi.Clear();
                cmbDnmYili.Text = "";
                cmbDnmAyi.Text = "";
            }
            catch (Exception q)
            {   
                 MessageBox.Show(q.Message);
            }
        }

        private void btnGunYeni_Click(object sender, EventArgs e)
        {
            try
            {
                cmbGun.Enabled = true;
                cmbGunDnmAdi.Enabled = true;
                btnGun_Ekle.Enabled = true;
                btnGun_Guncelle.Enabled = false;
                btnGunSil.Enabled = false;
                cmbGun.Text = "";
                cmbGunDnmAdi.Text = "";
                comboboxvericek();
            }
            catch (Exception q)
            {       
                 MessageBox.Show(q.Message);
            }
        }
        private void dgwGunListele_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                gunbilgisiyukle();
                cmbGun.Enabled = true;
                cmbGunDnmAdi.Enabled = true;
                btnGun_Guncelle.Enabled = true;
                btnGunSil.Enabled = true;
                btnGun_Ekle.Enabled = false;
            }
            catch (Exception q)
            {
                 MessageBox.Show(q.Message);
            }
        }
        private void btnHareketYeni_Click(object sender, EventArgs e)
        {
            try
            {
              //  comboboxvericek();
                btnHareketKaydet.Enabled = true;
                btnHareketGuncelle.Enabled = false;
                btnHareketSil.Enabled = false;
                mtHareketCıkıs.Enabled = true;
                mtHareketGiris.Enabled = true;
                cmbHareketAdiSoyadi.Enabled = true;
                cmbHareketDnmAdi.Enabled = true;
                cmbHareketGun.Enabled = true;
            }
            catch (Exception q)
            {
                
                 MessageBox.Show(q.Message);
            }
        }
        private void mesaisorgula()
        {
            try
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                SqlCommand cmdd = new SqlCommand("select 18000-ISNULL(SUM(DATEDIFF(MI,h.GirisSaati,h.CıkısSaati)),0) as 'Toplam' from Hareket h join dbo.Personel p on h.Personel_id=p.Personel_id where h.Personel_id=@persid ", baglanti);
                cmdd.Parameters.AddWithValue("@persid", cmbRaporlamaPersSec.SelectedValue.ToString());
                SqlDataReader dr = cmdd.ExecuteReader();
                string a;
                if (dr.Read())
                {
                    a = dr["Toplam"].ToString();
                    int b = Convert.ToInt32(a);
                    if (b < 0)
                    {
                        b *= -1;
                        string dosyaadi;
                        dosyaadi = cmbRaporlamaPersSec.Text + " " + cmbRaporlamaDonemSec.Text;
                        StreamWriter sw = new StreamWriter("C:\\Listeler\\" + dosyaadi + ".txt");
                        sw.Write("Sayın " + cmbRaporlamaPersSec.Text + " " + cmbRaporlamaDonemSec.Text + " toplamda yapmış olduğunuz mesai süreniz " + b + " dakikadır.");
                        sw.Close();
                    }
                    else
                    {
                        string dosyaadi;
                        dosyaadi = cmbRaporlamaPersSec.Text + " " + cmbRaporlamaDonemSec.Text;
                        StreamWriter sw = new StreamWriter(System.Windows.Forms.Application.StartupPath+"\\listeler\\"+ dosyaadi + ".txt");
                        sw.Write("Sayın " + cmbRaporlamaPersSec.Text + " " + cmbRaporlamaDonemSec.Text + " dönemine ait kalan çalışma süreniz " + a + " dakikadır.");
                        sw.Close();
                    }
                }
                dr.Close();
                baglanti.Close();
            }
            catch (Exception q)
            {    
                 MessageBox.Show(q.Message);
            }
        }
        private void btnRaporListele_Click(object sender, EventArgs e)
        {
            try
            {
                btnRaporExceleGonder.Enabled = true;
                btnRaporEpostaGonder.Enabled = true;
                
                    SqlCommand cmd = new SqlCommand("select p.adi,p.soyadi,g.Gun,d.DonemAyi,d.DonemYili,h.GirisSaati,h.CıkısSaati from dbo.Personel as p join dbo.Hareket as h  on p.Personel_id=h.Personel_id join dbo.Donem_gun g on h.gun_id=g.gun_id join dbo.Donem d on g.Donem_id=d.Donem_id where d.Donem_id=@dnmid and p.Personel_id=@prsid", baglanti);
                    cmd.Parameters.AddWithValue("@dnmid", cmbRaporlamaDonemSec.SelectedValue.ToString());
                       cmd.Parameters.AddWithValue("@prsid", cmbRaporlamaPersSec.SelectedValue.ToString()); 
                    SqlDataAdapter adp = new SqlDataAdapter(cmd);
                    System.Data.DataTable dtable = new System.Data.DataTable();
                    adp.Fill(dtable);
                    dgwRaporlama.DataSource = dtable;
                    baglanti.Close();
                   
            }
            catch (Exception q)
            {
                MessageBox.Show(q.Message);
            }
        }
        private void cmbRaporlamaDonemSec_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        private void cmbRaporlamaPersSec_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }
        private void btnRaporExceleGonder_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                Workbook wb = Excel.Workbooks.Add(XlSheetType.xlWorksheet);
                Worksheet ws = (Worksheet)Excel.ActiveSheet;
                Excel.Visible = true;
                ws.Cells[1, 1] = "Adı";
                ws.Cells[1, 2] = "Soyadı";
                ws.Cells[1, 3] = "Gün";
                ws.Cells[1, 4] = "Ay";
                ws.Cells[1, 5] = "Yıl";
                ws.Cells[1, 6] = "Giriş Saati";
                ws.Cells[1, 7] = "Çıkış Saati";
                for (int i = 2; i <= dgwRaporlama.Rows.Count + 1; i++)
                {
                    for (int j = 1; j <= 7; j++)
                    {
                        ws.Cells[i, j] = dgwRaporlama.Rows[i - 2].Cells[j - 1].Value;
                    }
                }
                }
            catch (Exception q)
            {   
                MessageBox.Show(q.Message);
            }            
        }
        private void btnRaporEpostaGonder_Click(object sender, EventArgs e)
        {
            try
            {
                mesaisorgula();
                OpenFileDialog da = new OpenFileDialog();
                da.Title = " Dosya Seç ";
                da.InitialDirectory = System.Windows.Forms.Application.StartupPath + "\\listeler";
                if (da.ShowDialog() == DialogResult.OK)
                {
                    if (baglanti.State == ConnectionState.Closed)
                    {
                        baglanti.Open();
                    }
                    SqlCommand cmdd = new SqlCommand("select p.E_Posta from Personel p where p.Personel_id=@persid ", baglanti);
                    cmdd.Parameters.AddWithValue("@persid", cmbRaporlamaPersSec.SelectedValue.ToString());
                    SqlDataReader dr = cmdd.ExecuteReader();
                    string a;
                    if (dr.Read())
                    {
                        a = dr["E_Posta"].ToString();
                        System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                        System.Net.NetworkCredential cred = new System.Net.NetworkCredential("pmys.docuart@gmail.com", "docu.12345");
                        mail.To.Add(a);
                        mail.Subject = "Çalışma ve Mesai Süreleri";
                        mail.From = new System.Net.Mail.MailAddress(a, "Docuart PMYS Bilgilendirme Sistemi");
                        mail.IsBodyHtml = true;  // mail.Body = " us:  pmys.docuart@gmail.com pw: docu.12345";
                        mail.Attachments.Add(new Attachment(da.FileName));
               System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587); smtp.UseDefaultCredentials = false;
                        smtp.EnableSsl = true;
                        smtp.Credentials = cred;
                        smtp.Send(mail);
                        LbLMailBilgi.Text = cmbRaporlamaPersSec.Text + "'a(e) Mail Gönderildi. ";
                    }
                    dr.Close();
                }
                else
                {
                    MessageBox.Show("E posta gönderirken önce dosya eki secilmelidir.");
                }
            }
            catch (Exception q )
            {
                MessageBox.Show(q.Message);
            }
        }


        int b=0;
        private void cmbGunDnmAdi_SelectedIndexChanged(object sender, EventArgs e)
        {
            b++;
            string df="";
            int t=0;
            string c="";
            if (b > 2)
            {
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                SqlCommand cmdd = new SqlCommand("Select d.DonemYili,d.DonemAyi,d.Donem_id from Donem d where d.Donem_id=@dnmid ", baglanti);
                cmdd.Parameters.AddWithValue("@dnmid", cmbGunDnmAdi.SelectedValue.ToString());
                SqlDataReader dr = cmdd.ExecuteReader();
               
                cmbHareketGun.Items.Clear();
                while (dr.Read())
                {
                    df = dr["DonemYili"].ToString();
                    c= dr["DonemAyi"].ToString();
                    
                }
                dr.Close();
         //       MessageBox.Show(df);
                if (c == "Ocak") t = 01;
                if (c == "Şubat") t = 02;
                if (c == "Mart") t = 03;
                if (c == "Nisan") t = 04;
                if (c == "Mayıs") t = 05;
                if (c == "Haziran") t = 06;
                if (c == "Temmuz") t = 07;
                if (c == "Ağustos") t = 08;
                if (c == "Eylül") t = 09;
                if (c == "Ekim") t = 10;
                if (c == "Kasım") t = 11;
                if (c == "Aralık") t = 12;
           //     MessageBox.Show(t.ToString());

                int gun;
                gun = DateTime.DaysInMonth(Convert.ToInt32(df), Convert.ToInt32(t));
             //   MessageBox.Show(gun.ToString());
                cmbGun.Items.Clear();
                for (int yaz = 1; yaz <= gun; yaz++)
                {
                            
                    cmbGun.Items.Add(yaz);
                }


            }
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {

                // cmbHareketGun.Items.Clear();
                comboboxvericek();

                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                SqlCommand cmdd = new SqlCommand("select g.gun from donem_gun g where g.donem_id=@dnmid ", baglanti);
                cmdd.Parameters.AddWithValue("@dnmid", cmbHareketDnmAdi.SelectedValue.ToString());
                //  MessageBox.Show(cmbHareketDnmAdi.SelectedValue.ToString());
                SqlDataReader dr = cmdd.ExecuteReader();
                string a;
                cmbHareketGun.Items.Clear();
                while (dr.Read())
                {
                    a = dr["gun"].ToString();
                    cmbHareketGun.Items.Add(a);
                }
                dr.Close();
            }
            catch (Exception)
            {
    
            }

    
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void cmbHareketDnmAdi_DisplayMemberChanged(object sender, EventArgs e)
        {
            
            
        }

        private void cmbHareketDnmAdi_ValueMemberChanged(object sender, EventArgs e)
        {

        }
        Excel.Application ExcelUygulama;
        Excel.Workbook ExcelProje;
        Excel.Worksheet ExcelSayfa;
        object Missing = System.Reflection.Missing.Value;
        Excel.Range ExcelRange;
        int rowCnt = 0;
        int columnCnt = 0;

        String dnmid; string kelime = "";
        private void button1_Click_1(object sender, EventArgs e)
        {
            DialogResult a = MessageBox.Show("Lütfen Dosyadaki Personel ile Sistemdeki Personelleri Kontrol Ediniz.\n Devam Ederseniz Sistemde Kayıtlı Olmayan Personeller Listelenecektir ve O Personellerin Veri Girişlerini Manuel Yapmanız Gerekmektedir.\n Personel Ekleyip Otomatik Kaydetmeyi Seçerseniz Çift Kayıt Sorunu İle Karşılaşacaksınız. Devam Etmek İster Misiniz? ","Uyarı",MessageBoxButtons.YesNo,MessageBoxIcon.Exclamation);
            if (DialogResult.Yes == a)
            {
                try
                {
                    listBox1.Items.Clear();
                    listBox2.Items.Clear();
                    listBox3.Items.Clear();

                    #region ExcelDosyaOkuma
                    DialogResult result = openFileDialog1.ShowDialog();
                    ExcelUygulama = new Excel.Application();
                    ExcelProje = ExcelUygulama.Workbooks.Open(openFileDialog1.FileName);
                    ExcelSayfa = (Excel.Worksheet)ExcelProje.Worksheets.get_Item(1);
                    ExcelRange = ExcelSayfa.UsedRange;
                    ExcelSayfa = (Excel.Worksheet)ExcelUygulama.ActiveSheet;

                    ExcelUygulama.Visible = false;
                    ExcelUygulama.AlertBeforeOverwriting = false;
                    rowCnt = ExcelRange.Rows.Count + 1;
                    columnCnt = ExcelRange.Columns.Count + 1;


                    String sdbconnection = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + openFileDialog1.FileName + "; Extended Properties='Excel 12.0; TypeGuessRows=0; HDR=YES; IMEX=1'";
                    OleDbConnection dbconnection = new OleDbConnection(sdbconnection);
                    dbconnection.Open();
                    OleDbDataAdapter dbadapter = new OleDbDataAdapter("Select * from [Page 1$]", dbconnection);
                    System.Data.DataTable dtable = new System.Data.DataTable();
                    dbadapter.Fill(dtable);
                    dataGridView1.DataSource = dtable;

                    string metn = "";
                    string veri;
                    #region ExcelSayfasınınİlkYarısı
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {

                        for (int j = 1; j <= 9; j++)
                        {

                            object hucre = ExcelSayfa.Cells[i, j];
                            Excel.Range bolge = ExcelSayfa.get_Range(hucre, hucre);
                            if (bolge.Value2 != null)
                            {

                                veri = bolge.Value2.ToString();
                                //     if (veri.Trim().Length == 16)
                                metn += " " + veri;
                            }

                        }
                        listBox1.Items.Add(metn.ToString());
                        metn = "";
                    }
                    #endregion
                    #region ExcelSayfasınınİkinciYarısı
                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {

                        for (int j = 10; j <= dataGridView1.ColumnCount; j++)
                        {

                            object hucre = ExcelSayfa.Cells[i, j];
                            Excel.Range bolge = ExcelSayfa.get_Range(hucre, hucre);
                            if (bolge.Value2 != null)
                            {

                                veri = bolge.Value2.ToString();
                                //    if (veri.Trim().Length == 16)

                                metn += " " + veri;
                            }

                        }
                        listBox1.Items.Add(metn.ToString());
                        metn = "";
                    }
                    #endregion
                    #endregion
                    #region Listbox1deki eleman sayısı
                    int kactane = listBox1.Items.Count;
                    // MessageBox.Show(kactane.ToString());
                    for (int q = 0; q < 226; q++)
                    {
                        string deneme = listBox1.Items[q].ToString();
                        if (deneme.Length < 5)
                        {
                            listBox1.Items.Remove(listBox1.Items[q]);
                        }
                    }
                    #endregion


                    KimlikKarsilastir();
                    MessageBox.Show("Aramalar Tamamlandı.");
                }
                catch (Exception q)
                {

                    MessageBox.Show(q.Message); this.Close();
                }
            }
            else
            {
                MessageBox.Show("Lütfen Personelleri Ekledikten Sonra Tekrar Deneyiniz.");
            }
        }
        ArrayList arraysayacdizisi = new ArrayList();
        ArrayList arrayforbitisi = new ArrayList();
        private void KimlikKarsilastir()
        {
            #region Rapor Donemini Alma
            String rapordonemivarmı = "";

            String RaporDonemi = "";
            foreach (String RaporAra in listBox1.Items)
            {
                if (RaporAra.Contains("Rapor Dönemi ") == true)
                {
                    RaporDonemi = RaporAra.ToString();
                    break;
                }
            }
            int donemuzunluk = RaporDonemi.Length;
            int charkopya = donemuzunluk - 14;
            String yeniDonem = RaporDonemi.Substring(14, charkopya); 
            #endregion
          //  MessageBox.Show(yeniDonem);

          
            #region DonemSorgulama

            if (baglanti.State == ConnectionState.Closed)
            {
                baglanti.Open();
            }
            SqlCommand dnm = new SqlCommand("select d.DonemAyi+' '+d.DonemYili as Donemler from donem d", baglanti); SqlDataReader dtr = dnm.ExecuteReader();
  String donembykharf;
            String dnmler;
            while (dtr.Read())
            {
                dnmler = dtr["Donemler"].ToString();
                donembykharf = dnmler.ToUpper();
                //  MessageBox.Show(donembykharf);
                if (yeniDonem == donembykharf)
                {
                     dtr.Close();
                    #region DonemidBulma
   string dnmayi = yeniDonem.Substring(0, yeniDonem.Length - 5);
   string dnmyili = yeniDonem.Substring(yeniDonem.Length - 4, 4);
   SqlCommand dnmidgetir = new SqlCommand("select Donem_id from Donem where DonemAyi='" + dnmayi + "' and DonemYili=" + dnmyili, baglanti); SqlDataReader ddtr = dnmidgetir.ExecuteReader();
   while (ddtr.Read())
   {
       dnmid = ddtr["Donem_id"].ToString();
      // MessageBox.Show(dnmid);
   }
   ddtr.Close();
                    #endregion
                    rapordonemivarmı = "true";
                    break;
                }
            }
         
            #endregion

            if (rapordonemivarmı == "true")
            {

                #region Kimlikverisinistringdenayirma
                int sayac = 0; int sira = 0;

                foreach (String gez in listBox1.Items)
                {
                    sayac++;
                    if (gez.Contains("TC Kimlik No:") == true)
                    {
                        listBox2.Items.Add((sira) + "-" + gez);
                        arraysayacdizisi.Add(sayac);
                        sira++;
                    }
                }
                int for_icin = 0;
                foreach (String forbitisnoktasi in listBox1.Items)
                {
                    for_icin++;
                    if (forbitisnoktasi.Contains("Dönemdeki Toplam Çalışma Süresi ") == true)
                    {
                        arrayforbitisi.Add(for_icin);
                    }
                }

                foreach (String item in listBox2.Items)
                {
                    int uzunlk = item.Length;
                    // MessageBox.Show(item.Substring(uzunlk - 11, 11));
                }
                #endregion

                #region VeritabanindanTcNoVerisiÇekme
                if (baglanti.State == ConnectionState.Closed)
                {
                    baglanti.Open();
                }
                SqlCommand cmdd = new SqlCommand("Select Tc_kimlik_no from Personel", baglanti); SqlDataReader dr = cmdd.ExecuteReader();
                String a;
                while (dr.Read())
                {
                    a = dr["Tc_kimlik_no"].ToString();
                    listBox3.Items.Add(a);
                }
                dr.Close();
                #endregion

               


                
                #region Veritabanikarsılastırma Ve Veritabanına Veri Girişi
                int dizisayaci = 0;
                foreach (String veritabanikimlik in listBox3.Items)
                {
                    foreach (String excelkmlk in listBox2.Items)
                    {
                        int uzunlk = excelkmlk.Length;
                        kelime = excelkmlk.Substring(uzunlk - 11, 11);
                        string index;
                        if (veritabanikimlik == kelime) //Kimlik Numarasının Eşleştiği Kısım ...
                        {
                            MessageBox.Show(veritabanikimlik + " Numaralı Kimlik Veritabanında Eşleşti");

                            #region Personel_id'yi bulan kısım
                            if (baglanti.State == ConnectionState.Closed)
                            {
                                baglanti.Open();
                            }
                            SqlCommand Kimlige_ait_kisiyi_bul = new SqlCommand("select p.Personel_id from Personel p  where p.Tc_kimlik_no=" + veritabanikimlik, baglanti); SqlDataReader kisidatareader = Kimlige_ait_kisiyi_bul.ExecuteReader();
                            string Veritabani_personel_id = "";
                            while (kisidatareader.Read())
                            {
                                Veritabani_personel_id = kisidatareader["Personel_id"].ToString();
                                //  MessageBox.Show(Veritabani_personel_id);
                            }
                            kisidatareader.Close();
                            #endregion
                            #region Günid ve parcalama
                            index = excelkmlk.Substring(0, (excelkmlk.Length) - 27);
                            dizisayaci = Convert.ToInt32(index);
                            string secilen, parcagun;
                            string girissaati, gelengunid = "";
                            string cikissaati;
                            int baslangic = Convert.ToInt32(arraysayacdizisi[dizisayaci]);
                            int bitis = Convert.ToInt32(arrayforbitisi[dizisayaci]);
                            for (int i = baslangic + 1; i <= bitis - 2; i++)
                            {
                                // MessageBox.Show(listBox1.Items[i].ToString());
                                secilen = listBox1.Items[i].ToString();
                                parcagun = secilen.Substring(1, 2);
                                girissaati = secilen.Substring(23, 5);
                                cikissaati = secilen.Substring(40, 5);
                                if (baglanti.State == ConnectionState.Closed)
                                {
                                    baglanti.Open();
                                }
                                SqlCommand gunidbul = new SqlCommand("select g.Gun_id from Donem_Gun g where g.Gun=" + parcagun + " and g.Donem_id =" + dnmid, baglanti); SqlDataReader sqldr = gunidbul.ExecuteReader();
                                while (sqldr.Read())
                                {
                                    gelengunid = sqldr["Gun_id"].ToString();
                                    //   MessageBox.Show(gelengunid);
                                }
                                sqldr.Close();
                            #endregion
                                #region VeritabanıEklemeKısmı
                                try
                                {
                                    SqlCommand cmd = new SqlCommand("INSERT INTO Hareket (GirisSaati,CıkısSaati,Personel_id,Gun_id) VALUES (@grs,@cks,@persid,@gunid)", baglanti);
                                    cmd.Parameters.AddWithValue("@grs", girissaati);
                                    cmd.Parameters.AddWithValue("@cks", cikissaati);
                                    cmd.Parameters.AddWithValue("@persid", Veritabani_personel_id);
                                    cmd.Parameters.AddWithValue("@gunid", gelengunid);
                                    if (baglanti.State == ConnectionState.Closed)
                                    {
                                        baglanti.Open();
                                    }
                                    cmd.ExecuteNonQuery();
                                    baglanti.Close();
                                    HareketBilgileriniGetir();
                                }
                                catch (Exception q)
                                {

                                    MessageBox.Show("Eklemede Hata var   : " + q.Message);
                                }
                                #endregion


                                //MessageBox.Show("Gun : "+parcagun);
                                //MessageBox.Show("GirisSaati: "+girissaati);
                                //MessageBox.Show("CıkışSaati: "+cikissaati);
                            }
                        }
                        else 
                        {
                            #region Lİstbx4_filtreleme
                            int list2count = listBox2.Items.Count;
                            int list3count = listBox3.Items.Count;
                            int fark = list2count - list3count;
                            if (listBox4.Items.Count <= fark - 1)
                            {
                                listBox4.Items.Add(kelime);
                            } 
                            #endregion
                        }

                    }

                }  
            }
              #endregion
            else
            {
                MessageBox.Show("Eklemeye Çalıştığınız Rapor Dönemi Bulunamadı.\n Lütfen Dönem Sekmesinden İlgili Dönemi Ekleyiniz.");
            }

        }
        private void Personel_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                ExcelProje.Close();
                ExcelUygulama.Quit();
            }
            catch (Exception)
            {
                System.Windows.Forms.Application.Exit();

            }
            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcesses())
            {
                if (p.ProcessName == "EXCEL")
                {
                    p.Kill();
                }
            }
        } 
    }
}