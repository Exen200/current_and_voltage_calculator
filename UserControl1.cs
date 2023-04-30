using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Numerics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Reflection.Metadata;
using System.Drawing.Imaging;
using System.CodeDom;
using System.Data.Common;
using Microsoft.VisualBasic.Devices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Reflection;

namespace WinFormsApp1
{
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }
        
        private void UserControl1_Load(object sender, EventArgs e)
        {
            
            dGrid1.RowCount = 1; // nadanie ilości wierszy w datagrid
            dGrid1.ColumnCount = 7; // nadanie ilości kolumn w datagrid
            dGrid1.Columns[0].HeaderCell.Value = "Nr odcinka"; // przypisanie nazwy do nagłówka danej kolumny
            dGrid1.Columns[1].HeaderCell.Value = "Dł.odcinka [km]"; // -||-
            dGrid1.Columns[2].HeaderCell.Value = "Prz. kabla/linii [mm^2]";
            dGrid1.Columns[3].HeaderCell.Value = "X'[Ω/km]";
            dGrid1.Columns[4].HeaderCell.Value = "Nr odbioru";
            dGrid1.Columns[5].HeaderCell.Value = "Moc [kW]";
            dGrid1.Columns[6].HeaderCell.Value = "Wsp. mocy [-]";
            
            DataGridViewComboBoxColumn column = new DataGridViewComboBoxColumn(); // stworzenie nowej kolumny o właściwościach comboboxa
            column.HeaderText = "Ch-ter obciążenia"; // przypisanie nazwy do nagłówka nowej kolumny
            column.Name = " "; // przypisanie nazwy nowej kolumnie
            column.Items.AddRange("Indukcyjny", "Pojemnościowy"); // dodanie wartości do nowej kolumny typu comboBox
            dGrid1.Columns.Add(column); // dodanie tej kolumny do datagrid

            dGrid1[0, 0].Value = 1; //nadanie wartości określonej komórce
            dGrid1[4, 0].Value = 1; //nadanie wartości określonej komórce
            dGrid1[0, 0].ReadOnly = true; //komórka readonly, czyli nie można jej edytować
            dGrid1[4, 0].ReadOnly = true;

            dGrid1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;  // blokowanie sortowania, żeby po kliknięciu w nagłówek nie zmieniało
            dGrid1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;  // się ustawienie danych tabeli, 
            dGrid1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;  // np. z 1 2 3 4 na 4 3 2 1, widać to po numerach
            dGrid1.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid1.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid1.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid1.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid1.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;

            dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView2.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView3.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView4.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView5.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;

            dataGridView1.AllowUserToResizeColumns = true;
        }
       
        
        private void dGrid1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            int  nrlini = (int)numericUpDown1.Value;  //ilość odcinków
            dGrid1.RowCount = nrlini;         // zmiana ilości odcinków
            
            for (int i = 0; i < nrlini; i++) //pętla do tworzenia numeracji w niżej podanych komórkach
            {
                dGrid1[0, i].Value = i + 1; //wartość w komórce [0,i] jest równe i + 1 i po każdej pętli zmnienia się
                dGrid1[4, i].Value = i + 1; // -||-, komórka oraz jej wartość

                dGrid1[0, i].ReadOnly = true;  //komórki z numeracją są readonly, więc nie można ich edytować
                dGrid1[4, i].ReadOnly = true;

            }

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            
            if (checkBox1.Checked) //jezeli checkbox zostanie klikniety
            {
                
                label8.Visible = true; // pojawia sie label - wprowadz dane transformatora
                label9.Visible = true;
                label10.Visible = true;
                label11.Visible = true;
                label12.Visible = true;
                label13.Visible = true;
                label14.Visible = true;
                label15.Visible = true;
                label16.Visible = true;
                label17.Visible = true;
                label18.Visible = true;
                textBox4.Visible = true;
                textBox5.Visible = true;
                textBox6.Visible = true;
                textBox7.Visible = true;
                textBox8.Visible = true;

                dGrid1[1, 0].ReadOnly = true; // długość odcinka w tej komórce -> readonly 
                dGrid1[2, 0].ReadOnly = true;  // przekroj odcinka w tej komórce -> readonly
                dGrid1[3, 0].ReadOnly = true;  // reaktancjajedn odcinka w tej komórce -> readonly

                dGrid1[1, 0].Value = null; // wartość tych komórek jest zmieniana na null, czyli brak wartości
                dGrid1[2, 0].Value = null; 
                dGrid1[3, 0].Value = null;
            }
            else
            {
                label8.Visible = false; // pojawia sie button - wprowadz dane transformatora
                label9.Visible = false;
                label10.Visible = false;
                label11.Visible = false;
                label12.Visible = false;
                label13.Visible = false;
                label14.Visible = false;
                label15.Visible = false;
                label16.Visible = false;
                label17.Visible = false;
                label18.Visible = false;
                textBox4.Visible = false;
                textBox5.Visible = false;
                textBox6.Visible = false;
                textBox7.Visible = false;
                textBox8.Visible = false;

                textBox4.Text = null;
                textBox5.Text = null;
                textBox6.Text = null;
                textBox7.Text = null;
                textBox8.Text = null;
               

                dGrid1[1, 0].ReadOnly = false; 
                dGrid1[2, 0].ReadOnly = false;  
                dGrid1[3, 0].ReadOnly = false;  

                dGrid1[1, 0].Value = null; // wartość tych komórek jest zmieniana na null, czyli brak wartości
                dGrid1[2, 0].Value = null;
                dGrid1[3, 0].Value = null;
            }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.RowHeadersWidth = 70;
            dataGridView2.RowHeadersWidth = 70;
            dataGridView3.RowHeadersWidth = 70;
            dataGridView4.RowHeadersWidth = 70;
            dataGridView5.RowHeadersWidth = 70;
            dataGridView1.AllowUserToResizeColumns = true;
            if (!checkBox1.Checked)
            {
                dataGridView5.Visible = false;
            }
            int nrlini = (int)numericUpDown1.Value;
            
            bool brakkolumn17 = false;
            bool braktxtboxa = false;
            bool txtboxzlyformat = false;
            bool dgridzlyformat = false;
            bool ujemnelubzero = false;
            if (!checkBox1.Checked)
            {
                for (int kolumna = 0; kolumna < 8; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlini; wiersz++)
                    {
                        if (dGrid1[kolumna, wiersz].Value == null)
                        {
                            brakkolumn17 = true;
                        }
                    }
                }
                for (int kolumna = 0; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlini; wiersz++)
                    {
                        if (dGrid1[kolumna, wiersz].Value != null)
                        {
                            string? value = dGrid1.Rows[wiersz].Cells[kolumna].Value.ToString();
                            decimal number;
                            if (Decimal.TryParse(value, out number))
                            {

                            }
                            else
                            {
                                dgridzlyformat = true;
                            }
                        }

                    }
                }
                for (int kolumna = 0; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlini; wiersz++)
                    {
                        if (dgridzlyformat == false)
                        {
                            double a;
                            a = Convert.ToDouble(dGrid1[kolumna, wiersz].Value);

                            if (a <= 0)
                            {
                                ujemnelubzero = true;
                            }
                        }
                    }
                }
                
                if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "")
                {
                    braktxtboxa = true;
                }
                if (textBox1.Text == "0" || textBox2.Text == "0" || textBox3.Text == "0")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == ",0" || textBox2.Text == ",0" || textBox3.Text == ",0")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == "0," || textBox2.Text == "0," || textBox3.Text == "0,")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == "," || textBox2.Text == "," || textBox3.Text == ",")
                {
                    txtboxzlyformat = true;
                }
            }
            
            else if (checkBox1.Checked)
            {
                for (int kolumna = 0; kolumna < 5; kolumna++)
                {
                    for (int wiersz = 1; wiersz < nrlini; wiersz++)
                    {
                        if (dGrid1[kolumna, wiersz].Value == null)
                        {
                            brakkolumn17 = true;
                        }
                    }
                }
                for (int kolumna = 5; kolumna < 8; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlini; wiersz++)
                    {
                        if (dGrid1[kolumna, wiersz].Value == null)
                        {
                            brakkolumn17 = true;
                        }
                    }
                }
                for (int kolumna = 0; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlini; wiersz++)
                    {
                        if (dGrid1[kolumna, wiersz].Value != null)
                        {
                            string? value = dGrid1.Rows[wiersz].Cells[kolumna].Value.ToString();
                            decimal number;
                            if (Decimal.TryParse(value, out number))
                            {

                            }
                            else
                            {
                                dgridzlyformat = true;
                            }
                        }

                    }
                }
                for (int kolumna = 0; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 1; wiersz < nrlini; wiersz++)
                    {
                        if (dgridzlyformat == false)
                        {
                            double a;
                            a = Convert.ToDouble(dGrid1[kolumna, wiersz].Value);

                            if (a <= 0)
                            {
                                ujemnelubzero = true;
                            }
                        }
                    }
                }
                
                if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text==""
                    || textBox5.Text == "" || textBox6.Text=="" || textBox7.Text==""|| textBox8.Text=="")
                {
                    braktxtboxa = true;
                }
                if (textBox1.Text == "0" || textBox2.Text == "0" || textBox3.Text == "0" || textBox4.Text == "0"
                   || textBox5.Text == "0" || textBox6.Text == "0" || textBox7.Text == "0" || textBox8.Text == "0")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == "0," || textBox2.Text == "0," || textBox3.Text == "0," || textBox4.Text == "0,"
                   || textBox5.Text == "0," || textBox6.Text == "0," || textBox7.Text == "0," || textBox8.Text == "0,")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == ",0" || textBox2.Text == ",0" || textBox3.Text == ",0" || textBox4.Text == ",0"
                   || textBox5.Text == ",0" || textBox6.Text == ",0" || textBox7.Text == ",0" || textBox8.Text == ",0")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == "," || textBox2.Text == "," || textBox3.Text == "," || textBox4.Text == ","
                        || textBox5.Text == "," || textBox6.Text == "," || textBox7.Text == "," || textBox8.Text == ",")
                {
                    txtboxzlyformat = true;
                }
            }

            if (brakkolumn17 == true || braktxtboxa == true)
            {
                string message5 = "Brak wymaganych wartości, uzpupełnij puste komórki";
                MessageBox.Show(message5);
            }
            else if (txtboxzlyformat == true)
            {
                string message6 = "Zły format, proszę wprowadzić liczbę";
                MessageBox.Show(message6);
            }
            else if (dgridzlyformat==true)
            {
                string message7 = "Zły format, proszę wprowadzić liczbę";
                MessageBox.Show(message7);
            }
            else if (ujemnelubzero == true)
            {
                string message8 = "Zła wartość liczby, proszę podać liczbę większą od zera";
                MessageBox.Show(message8);
            }
            else
            {
                dataGridView1.RowCount = nrlini;
                dataGridView2.RowCount = nrlini;
                dataGridView3.RowCount = nrlini;
                dataGridView4.RowCount = nrlini + 1;
                dataGridView5.RowCount = 2;
                double[] dlodc = new double[nrlini];
                double[] przkabla = new double[nrlini];
                double[] reaktjedn = new double[nrlini];
                double[] moc = new double[nrlini];
                double[] wspmoc = new double[nrlini];
                string[] chtrobc = new string[nrlini];
                double[] napwpkt = new double[nrlini + 1];

                double[] sinus = new double[nrlini];
                double[] cosinus = new double[nrlini];

                Complex[] pradodb = new Complex[nrlini];
                Complex[] impedancjalini = new Complex[nrlini];
                Complex impedancjatrafo;
                Complex[] pradrozpl = new Complex[nrlini + 1];
                Complex[] spadeknap = new Complex[nrlini];
                //double napznam = double.Parse(textBox3.Text)*1000;
                double przewodnoscwlasc = Convert.ToDouble(textBox2.Text);
                double napznam = Convert.ToDouble(textBox3.Text) * 1000;
                double nappoczatkow = Convert.ToDouble(textBox1.Text) * 1000;

                
                for (int h = 0; h < nrlini; h++)
                {
                    dataGridView1.Rows[h].HeaderCell.Value = "I" + (h + 1);


                    dataGridView2.Rows[h].HeaderCell.Value = "I" + h + (h + 1);


                    dataGridView3.Rows[h].HeaderCell.Value = "ΔU" + h + (h + 1);



                }

                for (int h = 0; h < nrlini + 1; h++)
                {
                    dataGridView4.Rows[h].HeaderCell.Value = "U" + h;
                }

                dataGridView1.RowHeadersWidth = 70;
                dataGridView2.RowHeadersWidth = 70;
                dataGridView3.RowHeadersWidth = 70;
                dataGridView4.RowHeadersWidth = 70;
                dataGridView5.RowHeadersWidth = 70;
                dataGridView1.ReadOnly = true;
                dataGridView2.ReadOnly = true;
                dataGridView3.ReadOnly = true;
                dataGridView4.ReadOnly = true;
                dataGridView5.ReadOnly = true;
                

                for (int i = 0; i < nrlini; i++)
                {
                    dlodc[i] = Convert.ToDouble(dGrid1[1, i].Value);
                    przkabla[i] = Convert.ToDouble(dGrid1[2, i].Value);
                    reaktjedn[i] = Convert.ToDouble(dGrid1[3, i].Value);
                    moc[i] = Convert.ToDouble(dGrid1[5, i].Value) * 1000;

                    chtrobc[i] = (string)dGrid1[7, i].Value;
                }

                for (int h = 0; h < nrlini; h++)
                {

                    wspmoc[h] = Convert.ToDouble(dGrid1[6, h].Value);
                    cosinus[h] = wspmoc[h];
                    sinus[h] = Math.Round(Math.Sqrt(1 - Math.Pow(cosinus[h], 2)), 2);
                    pradodb[h].real = Math.Round((moc[h] / (Math.Sqrt(3) * napznam * cosinus[h])) * cosinus[h], 2);

                    if (chtrobc[h] == "Indukcyjny")
                    {
                        pradodb[h].imaginary = Math.Round((moc[h] / (Math.Sqrt(3) * napznam * cosinus[h])) * (sinus[h] * (-1)), 2);
                    }

                    else
                    {
                        pradodb[h].imaginary = Math.Round((moc[h] / (Math.Sqrt(3) * napznam * cosinus[h])) * sinus[h], 2);
                    }



                    dataGridView1.RowCount = nrlini;
                    dataGridView1[0, h].Value = pradodb[h].ToString();




                }

                for (int h = nrlini - 1; h > -1; h--)
                {
                    if (h == nrlini - 1)
                    {
                        pradrozpl[h + 1].real = 0;
                        pradrozpl[h + 1].imaginary = 0;
                    }
                    pradrozpl[h].real = Math.Round(pradodb[h].real + pradrozpl[h + 1].real, 2);
                    pradrozpl[h].imaginary = Math.Round(pradodb[h].imaginary + pradrozpl[h + 1].imaginary, 2);
                    dataGridView2.RowCount = nrlini;
                    


                }
                for (int h = 0; h < nrlini; h++)
                {
                    dataGridView2[0, h].Value = pradrozpl[h].ToString();
                }

                //double rezyst;
                //double reakt;
                for (int l = 0; l < nrlini; l++)
                {
                    impedancjalini[l].real = Math.Round(((1000 / (przewodnoscwlasc * przkabla[l])) * dlodc[l]), 2);
                    impedancjalini[l].imaginary = Math.Round(reaktjedn[l] * dlodc[l], 2);
                    if (checkBox1.Checked)
                    {
                        double deltapcutrafo = Convert.ToDouble(textBox4.Text);
                        double napzwarciatrafo = Convert.ToDouble(textBox5.Text);
                        double napuzwjgornego = Convert.ToDouble(textBox7.Text);
                        double napuzwdolnego = Convert.ToDouble(textBox8.Text);
                        double moctrafo = Convert.ToDouble(textBox6.Text);
                        impedancjatrafo.real = Math.Round((deltapcutrafo * Math.Pow(napuzwdolnego, 2)) / (100 * moctrafo), 6);
                        impedancjatrafo.imaginary = Math.Round((napzwarciatrafo * Math.Pow(napuzwdolnego, 2)) / (100 * moctrafo), 6);

                        impedancjalini[0].real = impedancjatrafo.real;
                        impedancjalini[0].imaginary = impedancjatrafo.imaginary;

                    }


                 
                }

                for (int c = 0; c < nrlini; c++)
                {
                    dataGridView3.RowCount = nrlini;
                    spadeknap[c] = pradrozpl[c] * impedancjalini[c];
                    spadeknap[c].real = Math.Round(spadeknap[c].real * Math.Sqrt(3), 2);
                    spadeknap[c].imaginary = Math.Round(spadeknap[c].imaginary * Math.Sqrt(3), 2);

                    dataGridView3[0, c].Value = spadeknap[c].real.ToString();


                }
                Complex pradwysoki = new Complex(0, 0);
                Complex pradrozplzas = new Complex(0, 0);
                if (checkBox1.Checked) 
                {

                    dataGridView5.Visible = true;

                    double napuzwjgornego = Convert.ToDouble(textBox7.Text);
                    double napuzwdolnego = Convert.ToDouble(textBox8.Text);
                    
                    pradrozplzas.real = pradrozpl[0].real;
                    pradrozplzas.imaginary = pradrozpl[0].imaginary;
                    pradwysoki = pradrozplzas * ( napuzwdolnego/ napuzwjgornego);
                    pradwysoki.real = Math.Round(pradwysoki.real,2);
                    pradwysoki.imaginary = Math.Round(pradwysoki.imaginary, 2);
                    dataGridView5.Rows[0].HeaderCell.Value = "Iz0" + "(" + napuzwdolnego + ")";
                    dataGridView5.Rows[1].HeaderCell.Value = "Iz0" + "(" + napuzwjgornego + ")";
                    dataGridView5[0, 0].Value = pradrozplzas;
                    dataGridView5[0, 1].Value = pradwysoki;
                }
                
                if (!checkBox1.Checked)
                {
                    napwpkt[0] = nappoczatkow;
                }
                else if (checkBox1.Checked)
                {
                    double napuzwjgornego = Convert.ToDouble(textBox7.Text);
                    double napuzwdolnego = Convert.ToDouble(textBox8.Text);
                    nappoczatkow = nappoczatkow * (napuzwdolnego / napuzwjgornego);
                    napwpkt[0] = nappoczatkow;
                    
                }
                for (int d = 1; d < nrlini + 1; d++)
                {

                    napwpkt[d] = napwpkt[d - 1] - spadeknap[d - 1].real;

                    


                }
                for (int d = 0; d < nrlini + 1; d++)
                {
                    napwpkt[d] = Math.Round(napwpkt[d], 2);

                    dataGridView4[0, d].Value = napwpkt[d].ToString();
                }

            }
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataGridViewComboBoxColumn column = new DataGridViewComboBoxColumn(); 
            column.HeaderText = "Ch-ter obciążenia"; 
            column.Name = " "; 
            column.Items.AddRange("Indukcyjny", "Pojemnościowy"); 
           
            numericUpDown1.Value = 1;
            for (int i=1;i<4;i++)
            {
                dGrid1[i, 0].Value = null;
            }
            for (int i = 5; i < 7; i++)
            {
                dGrid1[i, 0].Value = null;
            }
            dGrid1.Columns.RemoveAt(7);
            dGrid1.Columns.Add(column);
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            checkBox1.Checked = false;
            
            dataGridView1.RowCount = 1;
            dataGridView1[0, 0].Value = null;
            dataGridView2.RowCount = 1;
            dataGridView2[0, 0].Value = null;
            dataGridView3.RowCount = 1;
            dataGridView3[0, 0].Value = null;
            dataGridView4.RowCount = 1;
            dataGridView4[0, 0].Value = null;
            dataGridView5.RowCount = 1;
            dataGridView5[0, 0].Value = null;

            dataGridView1.Rows[0].HeaderCell.Value = " ";
            dataGridView2.Rows[0].HeaderCell.Value = " ";
            dataGridView3.Rows[0].HeaderCell.Value = " ";
            dataGridView4.Rows[0].HeaderCell.Value = " ";
            dataGridView5.Rows[0].HeaderCell.Value = " ";

            pictureBox1.Image = null;
            //pictureBox1.Invalidate();
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            
            int iloscodinkow = (int)numericUpDown1.Value;
             
            string odbnr;
            string path = Directory.GetCurrentDirectory();

            string fullpath = path + "\\" + "obrazki\\";
            if (!checkBox1.Checked)
            {
                odbnr = "j" + iloscodinkow;

                Image image = Image.FromFile(fullpath + odbnr +".jpg");
                this.pictureBox1.Image = image;
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            }
            else if (checkBox1.Checked)
            {
                odbnr = "j" + iloscodinkow + "t";
                Image image = Image.FromFile(fullpath + odbnr + ".jpg");
                this.pictureBox1.Image = image;
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            }
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image == null)
            { string message = "Brak układu, załaduj układ";
                MessageBox.Show(message);
            }
            else
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "JPeg Image|*.jpg|PNG Image|*.png";
                saveFileDialog1.Title = "Save an Image File";
                saveFileDialog1.ShowDialog();

                if (saveFileDialog1.FileName != "")
                {
                    System.IO.FileStream fs = (System.IO.FileStream)saveFileDialog1.OpenFile();
                   
                    switch (saveFileDialog1.FilterIndex)
                    {
                        case 1:
                            this.pictureBox1.Image.Save(fs,
                              System.Drawing.Imaging.ImageFormat.Jpeg);
                            break;
                        case 2:
                            this.pictureBox1.Image.Save(fs,
                              System.Drawing.Imaging.ImageFormat.Png);
                            break;
                    }

                    fs.Close();
                }
            }
        
        }
        private void copyAlltoClipboard1()
        {
           dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
           dataGridView1.MultiSelect = true; 
           dataGridView1.SelectAll();
           DataObject dataObj1 = dataGridView1.GetClipboardContent();
           if (dataObj1 != null)
                Clipboard.SetDataObject(dataObj1);
            
        }
        private void copyAlltoClipboard2()
        {
           dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
           dataGridView2.MultiSelect = true;
           dataGridView2.SelectAll();
           DataObject dataObj2 = dataGridView2.GetClipboardContent();
           if (dataObj2 != null)
                Clipboard.SetDataObject(dataObj2);
           
        }
        private void copyAlltoClipboard3()
        {

            dataGridView3.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView3.MultiSelect = true;
            dataGridView3.SelectAll();
            DataObject dataObj3 = dataGridView3.GetClipboardContent();
            if (dataObj3 != null)
                Clipboard.SetDataObject(dataObj3);
            
        } 
        private void copyAlltoClipboard4()
        {
            dataGridView4.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView4.MultiSelect = true;
            dataGridView4.SelectAll();
            DataObject dataObj4 = dataGridView4.GetClipboardContent();
            if (dataObj4 != null)
                Clipboard.SetDataObject(dataObj4);
            
        }
        private void copyAlltoClipboard5()
        {
            dataGridView5.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView5.MultiSelect = true;
            dataGridView5.SelectAll();
            DataObject dataObj5 = dataGridView5.GetClipboardContent();
            if (dataObj5 != null)
                Clipboard.SetDataObject(dataObj5);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            if (dataGridView1.Rows[0].Cells[0].Value ==null)
            {
                string message1 = "Brak wartości w tabelach z wynikami, najpierw oblicz wyniki";
                MessageBox.Show(message1);
            }
            else 
            {
                copyAlltoClipboard1();
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                copyAlltoClipboard2();
                object misValue1 = System.Reflection.Missing.Value;
                //xlWorkBook = xlexcel.Workbooks.Add(misValue1);
                //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Excel.Range CR1 = (Excel.Range)xlWorkSheet.Cells[1, 5];
                CR1.Select();
                xlWorkSheet.PasteSpecial(CR1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                copyAlltoClipboard3();
                object misValue2 = System.Reflection.Missing.Value;
                Excel.Range CR2 = (Excel.Range)xlWorkSheet.Cells[1, 9];
                CR2.Select();
                xlWorkSheet.PasteSpecial(CR2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                copyAlltoClipboard4();
                object misValue3 = System.Reflection.Missing.Value;
                Excel.Range CR3 = (Excel.Range)xlWorkSheet.Cells[1, 13];
                CR3.Select();
                xlWorkSheet.PasteSpecial(CR3, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                
                if (checkBox1.Checked)
                {
                    copyAlltoClipboard5();
                    object misValue4 = System.Reflection.Missing.Value;
                    Excel.Range CR4 = (Excel.Range)xlWorkSheet.Cells[1, 17];
                    CR4.Select();
                    xlWorkSheet.PasteSpecial(CR4, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
                }
            }
           
        }

        private void dGrid1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            /*int nrlini = (int)numericUpDown1.Value;
            
            if (!checkBox1.Checked)
            {
               for (int kolumna = 1; kolumna < 7; kolumna++) 
               {
                   for (int wiersz=0;wiersz<nrlini;wiersz++)
                   {
                        if ( e.ColumnIndex == kolumna && e.RowIndex == wiersz)
                        {
                            decimal i;

                            if (!decimal.TryParse(Convert.ToString(e.FormattedValue), out i))
                            {
                                e.Cancel = true;
                                MessageBox.Show("Zły format, proszę wprowadzić liczbę");
                            }
                            else
                            {
                                // the input is numeric 
                            }
                        }
                   }
                    
               }
                   
            }
            else if (checkBox1.Checked)
            {
                for (int kolumna = 1; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 1; wiersz < nrlini; wiersz++)
                    {
                            if (e.ColumnIndex == kolumna && e.RowIndex == wiersz)
                            {
                                decimal i;

                                if (!decimal.TryParse(Convert.ToString(e.FormattedValue), out i))
                                {
                                    e.Cancel = true;
                                    MessageBox.Show("Zły format, proszę wprowadzić liczbę");
                                }
                                else
                                {
                                    // the input is numeric 
                                }

                            }
                    }
                }
                for (int kolumna = 5; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlini; wiersz++)
                    {

                        if (e.ColumnIndex == kolumna && e.RowIndex == wiersz)
                        {
                            decimal i;

                            if (!decimal.TryParse(Convert.ToString(e.FormattedValue), out i))
                            {
                                e.Cancel = true;
                                MessageBox.Show("Zły format, proszę wprowadzić liczbę");
                            }
                            else
                            {
                                // the input is numeric 
                            }

                        }


                    }

                }
            }*/
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar)
        && !char.IsDigit(e.KeyChar)
        && e.KeyChar != ',')
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if (e.KeyChar == ','
                && ((System.Windows.Forms.TextBox)sender).Text.IndexOf(',') > -1)
            {
                e.Handled = true;
            }
        
        
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar)
                    && !char.IsDigit(e.KeyChar)
                    && e.KeyChar != ',')
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if (e.KeyChar == ','
                && ((System.Windows.Forms.TextBox)sender).Text.IndexOf(',') > -1)
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar)
        && !char.IsDigit(e.KeyChar)
        && e.KeyChar != ',')
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if (e.KeyChar == ','
                && ((System.Windows.Forms.TextBox)sender).Text.IndexOf(',') > -1)
            {
                e.Handled = true;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (checkBox1.Checked)
            {
                if (!char.IsControl(e.KeyChar)
                        && !char.IsDigit(e.KeyChar)
                        && e.KeyChar != ',')
                {
                    e.Handled = true;
                }

                // only allow one decimal point
                if (e.KeyChar == ','
                    && ((System.Windows.Forms.TextBox)sender).Text.IndexOf(',') > -1)
                {
                    e.Handled = true;
                }
            }
            
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (checkBox1.Checked)
            {
                if (!char.IsControl(e.KeyChar)
                        && !char.IsDigit(e.KeyChar)
                        && e.KeyChar != ',')
                {
                    e.Handled = true;
                }

                // only allow one decimal point
                if (e.KeyChar == ','
                    && ((System.Windows.Forms.TextBox)sender).Text.IndexOf(',') > -1)
                {
                    e.Handled = true;
                }
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (checkBox1.Checked)
            {
                if (!char.IsControl(e.KeyChar)
                        && !char.IsDigit(e.KeyChar)
                        && e.KeyChar != ',')
                {
                    e.Handled = true;
                }

                // only allow one decimal point
                if (e.KeyChar == ','
                    && ((System.Windows.Forms.TextBox)sender).Text.IndexOf(',') > -1)
                {
                    e.Handled = true;
                }
            }
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (checkBox1.Checked)
            {
                if (!char.IsControl(e.KeyChar)
                        && !char.IsDigit(e.KeyChar)
                        && e.KeyChar != ',')
                {
                    e.Handled = true;
                }

                // only allow one decimal point
                if (e.KeyChar == ','
                    && ((System.Windows.Forms.TextBox)sender).Text.IndexOf(',') > -1)
                {
                    e.Handled = true;
                }
            }
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (checkBox1.Checked)
            {
                if (!char.IsControl(e.KeyChar)
                        && !char.IsDigit(e.KeyChar)
                        && e.KeyChar != ',')
                {
                    e.Handled = true;
                }

                // only allow one decimal point
                if (e.KeyChar == ','
                    && ((System.Windows.Forms.TextBox)sender).Text.IndexOf(',') > -1)
                {
                    e.Handled = true;
                }
            }
        }

        private void button_przyklad1_Click(object sender, EventArgs e)
        {

            string path = Directory.GetCurrentDirectory();
            string fullpath = path + "\\" + "przyklady\\przyklad1j.txt";
            string text = File.ReadAllText(fullpath);
           

            string[] words = text.Split(new char[] { ' ', '\n' });
            textBox1.Text = Convert.ToString(words[0]);
            textBox3.Text = Convert.ToString(words[1]);
            textBox2.Text = Convert.ToString(words[2]);

            string wspmoc1 = "Indukcyjny";
            string wspmoc2 = "Indukcyjny";
            string wspmoc3 = "Indukcyjny";

            int nrlini = Convert.ToInt32(words[3]);
            numericUpDown1.Value = nrlini;
            
            for (int kolumna = 0; kolumna < 7; kolumna++)
            {
                for (int wiersz = 0; wiersz < nrlini; wiersz++)
                {
                    dGrid1[kolumna, wiersz].Value = words[4 + wiersz * 7 + kolumna];


                }
            }
            dGrid1[7, 0].Value = wspmoc1;
            dGrid1[7, 1].Value = wspmoc2;
            dGrid1[7, 2].Value = wspmoc3;
        }

        private void button_przyklad2_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = true;
            string path1 = Directory.GetCurrentDirectory();
            string fullpath1 = path1 + "\\" + "przyklady\\przyklad2j.txt";
            string text1 = File.ReadAllText(fullpath1);
            

            string[] words1 = text1.Split(new char[] { ' ', '\n' });
            textBox1.Text = Convert.ToString(words1[0]);
            textBox3.Text = Convert.ToString(words1[1]);
            textBox2.Text = Convert.ToString(words1[2]);
            
            textBox4.Text = Convert.ToString(words1[4]);
            textBox5.Text = Convert.ToString(words1[5]);
            textBox6.Text = Convert.ToString(words1[6]);
            textBox7.Text = Convert.ToString(words1[7]);
            textBox8.Text = Convert.ToString(words1[8]);

            string wspmoc1 = "Indukcyjny";
            string wspmoc2 = "Indukcyjny";
            string wspmoc3 = "Indukcyjny";

            int nrlini = Convert.ToInt32(words1[3]);
            numericUpDown1.Value = nrlini;
            for (int kolumna = 0; kolumna < 7; kolumna++)
            {
                for (int wiersz = 1; wiersz < nrlini; wiersz++)
                {
                    dGrid1[kolumna, wiersz].Value = words1[2 + wiersz * 7 + kolumna];


                }
            }
            string ada= "0,8";
            dGrid1[7, 0].Value = wspmoc1;
            dGrid1[7, 1].Value = wspmoc2;
            dGrid1[7, 2].Value = wspmoc3;
            dGrid1[5, 0].Value = 1200;
            dGrid1[6, 0].Value = ada;
        }
    }
}
