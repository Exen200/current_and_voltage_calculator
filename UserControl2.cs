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
    public partial class UserControl2 : UserControl
    {
        public UserControl2()
        {
            InitializeComponent();
        }

        private void UserControl2_Load(object sender, EventArgs e)
        {
            dGrid1.RowCount = 2; // nadanie ilości wierszy w datagrid
            dGrid1.ColumnCount = 7; // nadanie ilości kolumn w datagrid
            dGrid1.Columns[0].HeaderCell.Value = "Nr odcinka"; // przypisanie nazwy do nagłówka danej kolumny
            dGrid1.Columns[1].HeaderCell.Value = "Dł.odcinka [km]";  // -||-
            dGrid1.Columns[2].HeaderCell.Value = "Prz. kabla/linii [mm^2]";
            dGrid1.Columns[3].HeaderCell.Value = "X'[Ω/km]";
            dGrid1.Columns[4].HeaderCell.Value = "Nr odbioru";
            dGrid1.Columns[5].HeaderCell.Value = "Moc [kW]";
            dGrid1.Columns[6].HeaderCell.Value = "Wsp. mocy [-]";
            DataGridViewComboBoxColumn column = new DataGridViewComboBoxColumn(); // stworzenie nowej kolumny o właściwościach comboboxa
            column.HeaderText = "Ch-ter obciążenia"; // przypisanie nazwy do nagłówka nowej kolumny
            column.Name = " "; // przypisanie nazwy nowej kolumny
            column.Items.AddRange("Indukcyjny", "Pojemnościowy"); // dodanie wartości do nowej kolumny typu comboBox
            dGrid1.Columns.Add(column); // dodanie tej kolumny do datagrid
            int? value = 0; //stworzenie zmiennej pod nulla

            if (value == 0) //sprawdzenie czy wartość jest równa 0
            {
                value = null; //przypisanie do niej nulla
            }
            dGrid1[0, 0].Value = 1;
            dGrid1[0, 1].Value = 2; //nadanie wartości określonej komórce; dGrid[nr kolumny,nr wiersza]
            dGrid1[4, 0].Value = value;
            dGrid1[4, 1].Value = 1;//przypisanie nulla do wartość w tej komórce (komórka z numerem odbioru)
            dGrid1[5, 0].Value = value;
            dGrid1[6, 0].Value = value;
            dGrid1[7, 0].Value = value;
            dGrid1[0, 0].ReadOnly = true; //komórka readonly, czyli nie można jej edytować
            dGrid1[0, 1].ReadOnly = true;
            dGrid1[4, 0].ReadOnly = true;
            dGrid1[4, 1].ReadOnly = true;
            dGrid1[5, 0].ReadOnly = true;
            dGrid1[6, 0].ReadOnly = true;
            dGrid1[7, 0].ReadOnly = true;

            dGrid1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;  // blokowanie sortowania, żeby po kliknięciu w nagłówek nie zmieniało
            dGrid1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;  // się ustawienie danych tabeli, 
            dGrid1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;  // np. z 1 2 3 4 na 4 3 2 1, widać to po numerach
            dGrid1.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid1.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid1.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid1.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid1.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;

            dataGridView11.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView22.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView33.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView44.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            int nrlini = (int)numericUpDown1.Value;  //ilość odcinków
            dGrid1.RowCount = nrlini;         // zmiana ilości odcinków

            for (int i = 0; i < nrlini; i++)  //pętla do tworzenia numeracji w niżej podanych komórkach
            {
                dGrid1[0, i].Value = i + 1;  //wartość w komórce [0,i] jest równe i + 1 i po każdej pętli zmnienia się
                dGrid1[4, i].Value = i;   // komórka oraz jej wartość

                int? value = 0;   //stworzenie zmiennej pod nulla

                if (value == 0)   //sprawdzenie czy wartość jest równa 0
                {
                    value = null; //przypisanie do niej nulla
                }

                dGrid1[4, 0].Value = value; //przypisanie nulla do wartość w tej komórce (komórka z numerem odbioru)
                dGrid1[0, i].ReadOnly = true; // 1 kolumny (czyli zerowej ) nie można edytować
                dGrid1[4, i].ReadOnly = true;

                if (checkBox1.Checked)
                {
                    dGrid1[1, nrlini - 1].ReadOnly = true;
                    dGrid1[2, nrlini - 1].ReadOnly = true;
                    dGrid1[3, nrlini - 1].ReadOnly = true;
                    dGrid1[1, nrlini - 2].ReadOnly = false;
                    dGrid1[2, nrlini - 2].ReadOnly = false;
                    dGrid1[3, nrlini - 2].ReadOnly = false;
                    dGrid1[1, 0].ReadOnly = true;
                    dGrid1[2, 0].ReadOnly = true;
                    dGrid1[3, 0].ReadOnly = true;
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            int indexc = (int)numericUpDown1.Value; // ilość wierszy w data grid, liczy 1,2,3,4, dlatego jest minus 1


            if (checkBox1.Checked)  //jezeli checkbox zostanie klikniety
            {

                label10.Visible = true; // label znika
                label11.Visible = true;
                label12.Visible = true;
                label13.Visible = true;
                label14.Visible = true;
                label15.Visible = true;
                label16.Visible = true;
                label17.Visible = true;
                label18.Visible = true;
                label19.Visible = true;
                label20.Visible = true;
                label21.Visible = true;
                label21.Visible = true;
                label22.Visible = true;
                label23.Visible = true;
                label24.Visible = true;
                label25.Visible = true;
                label26.Visible = true;
                label27.Visible = true;
                label28.Visible = true;
                label29.Visible = true;
                label30.Visible = true;
                label31.Visible = true;


                textBox7.Visible = true;
                textBox8.Visible = true;
                textBox9.Visible = true;
                textBox10.Visible = true;
                textBox11.Visible = true;
                textBox12.Visible = true;
                textBox13.Visible = true;
                textBox14.Visible = true;
                textBox15.Visible = true;
                textBox16.Visible = true;

                dGrid1[1, 0].ReadOnly = true; // długość odcinka w tej komórce -> readonly 
                dGrid1[2, 0].ReadOnly = true;  // przekroj odcinka w tej komórce -> readonly
                dGrid1[3, 0].ReadOnly = true;  // reaktancjajedn odcinka w tej komórce -> readonly

                dGrid1[1, dGrid1.RowCount - 1].ReadOnly = true;
                dGrid1[2, dGrid1.RowCount - 1].ReadOnly = true;
                dGrid1[3, dGrid1.RowCount - 1].ReadOnly = true;

                dGrid1[1, 0].Value = null;
                dGrid1[2, 0].Value = null;
                dGrid1[3, 0].Value = null;

                dGrid1[1, dGrid1.RowCount - 1].Value = null;
                dGrid1[2, dGrid1.RowCount - 1].Value = null;
                dGrid1[3, dGrid1.RowCount - 1].Value = null;


            }
            else
            {

                label10.Visible = false;// label znika
                label11.Visible = false;
                label12.Visible = false;
                label13.Visible = false;
                label14.Visible = false;
                label15.Visible = false;
                label16.Visible = false;
                label17.Visible = false;
                label18.Visible = false;
                label19.Visible = false;
                label20.Visible = false;
                label21.Visible = false;
                label21.Visible = false;
                label22.Visible = false;
                label23.Visible = false;
                label24.Visible = false;
                label25.Visible = false;
                label26.Visible = false;
                label27.Visible = false;
                label28.Visible = false;
                label29.Visible = false;
                label30.Visible = false;
                label31.Visible = false;


                textBox7.Visible = false;
                textBox8.Visible = false;
                textBox9.Visible = false;
                textBox10.Visible = false;
                textBox11.Visible = false;
                textBox12.Visible = false;
                textBox13.Visible = false;
                textBox14.Visible = false;
                textBox15.Visible = false;
                textBox16.Visible = false;

                textBox7.Text = null;
                textBox8.Text = null;
                textBox9.Text = null;
                textBox10.Text = null;
                textBox11.Text = null;
                textBox12.Text = null;
                textBox13.Text = null;
                textBox14.Text = null;
                textBox15.Text = null;
                textBox16.Text = null;


                dGrid1[1, 0].ReadOnly = false;
                dGrid1[2, 0].ReadOnly = false;
                dGrid1[3, 0].ReadOnly = false;

                dGrid1[1, dGrid1.RowCount - 1].ReadOnly = false;
                dGrid1[2, dGrid1.RowCount - 1].ReadOnly = false;
                dGrid1[3, dGrid1.RowCount - 1].ReadOnly = false;

                dGrid1[1, 0].Value = null;
                dGrid1[2, 0].Value = null;
                dGrid1[3, 0].Value = null;

                dGrid1[1, dGrid1.RowCount - 1].Value = null;
                dGrid1[2, dGrid1.RowCount - 1].Value = null;
                dGrid1[3, dGrid1.RowCount - 1].Value = null;


            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView11.RowHeadersWidth = 70;
            dataGridView22.RowHeadersWidth = 70;
            dataGridView33.RowHeadersWidth = 70;
            dataGridView44.RowHeadersWidth = 70;
            dataGridView1.RowHeadersWidth = 70;
            if (!checkBox1.Checked)
            {
                dataGridView1.Visible = false;
            }

            int nrlini = (int)numericUpDown1.Value;
            bool brakkolumn17 = false;
            bool braktxtboxa = false;
            bool txtboxzlyformat = false;
            bool dgridzlyformat = false;
            bool ujemnelubzero = false;
            if (!checkBox1.Checked)
            {
                for (int kolumna = 0; kolumna < 4; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlini; wiersz++)
                    {
                        if (dGrid1[kolumna, wiersz].Value == null)
                        {
                            brakkolumn17 = true;
                        }
                    }
                }
                for (int kolumna = 4; kolumna < 8; kolumna++)
                {
                    for (int wiersz = 1; wiersz < nrlini; wiersz++)
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
                for (int kolumna = 0; kolumna < 4; kolumna++)
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
                for (int kolumna = 4; kolumna < 7; kolumna++)
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

                if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == ""
                    || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "")
                {
                    braktxtboxa = true;
                }
                if (textBox1.Text == "0" || textBox3.Text == "0"
                    || textBox4.Text == "0" || textBox6.Text == "0")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == "0," || textBox3.Text == "0,"
                    || textBox4.Text == "0," || textBox6.Text == "0,")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == ",0" || textBox3.Text == ",0"
                    || textBox4.Text == ",0" || textBox6.Text == ",0")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == "," || textBox2.Text == "," || textBox3.Text == ","
                        || textBox4.Text == "," || textBox5.Text == "," || textBox6.Text == ",")
                {
                    txtboxzlyformat = true;
                }
            }


            else if (checkBox1.Checked)
            {
                for (int kolumna = 0; kolumna < 4; kolumna++)
                {
                    for (int wiersz = 1; wiersz < nrlini - 1; wiersz++)
                    {
                        if (dGrid1[kolumna, wiersz].Value == null)
                        {
                            brakkolumn17 = true;
                        }
                    }
                }
                for (int kolumna = 4; kolumna < 8; kolumna++)
                {
                    for (int wiersz = 1; wiersz < nrlini; wiersz++)
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
                for (int kolumna = 0; kolumna < 4; kolumna++)
                {
                    for (int wiersz = 1; wiersz < nrlini - 1; wiersz++)
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
                for (int kolumna = 4; kolumna < 7; kolumna++)
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

                if (textBox1.Text == " " || textBox2.Text == " " || textBox3.Text == " " || textBox4.Text == " "
                    || textBox5.Text == " " || textBox6.Text == " " || textBox7.Text == " " || textBox8.Text == " "
                    || textBox9.Text == " " || textBox10.Text == " " || textBox10.Text == " " || textBox11.Text == " "
                    || textBox12.Text == " " || textBox13.Text == " " || textBox14.Text == " " || textBox15.Text == " "
                    || textBox16.Text == " ")
                {
                    braktxtboxa = true;
                }
                if (textBox1.Text == "0" || textBox3.Text == "0" || textBox4.Text == "0"
                    || textBox6.Text == "0" || textBox7.Text == "0" || textBox8.Text == "0"
                    || textBox9.Text == "0" || textBox10.Text == "0" || textBox10.Text == "0" || textBox11.Text == "0"
                    || textBox12.Text == "0" || textBox13.Text == "0" || textBox14.Text == "0" || textBox15.Text == "0"
                    || textBox16.Text == "0")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == "0," || textBox3.Text == "0," || textBox4.Text == "0,"
                     || textBox6.Text == "0," || textBox7.Text == "0," || textBox8.Text == "0,"
                    || textBox9.Text == "0," || textBox10.Text == "0," || textBox10.Text == "0," || textBox11.Text == "0,"
                    || textBox12.Text == "0," || textBox13.Text == "0," || textBox14.Text == "0," || textBox15.Text == "0,"
                    || textBox16.Text == "0,")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == ",0" || textBox3.Text == ",0" || textBox4.Text == ",0"
                     || textBox6.Text == ",0" || textBox7.Text == ",0" || textBox8.Text == ",0"
                    || textBox9.Text == ",0" || textBox10.Text == ",0" || textBox10.Text == ",0" || textBox11.Text == ",0"
                    || textBox12.Text == ",0" || textBox13.Text == ",0" || textBox14.Text == ",0" || textBox15.Text == ",0"
                    || textBox16.Text == ",0")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == "," || textBox2.Text == "," || textBox3.Text == "," || textBox4.Text == ","
                    || textBox5.Text == "," || textBox6.Text == "," || textBox7.Text == "," || textBox8.Text == ","
                    || textBox9.Text == "," || textBox10.Text == "," || textBox10.Text == "," || textBox11.Text == ","
                    || textBox12.Text == "," || textBox13.Text == "," || textBox14.Text == "," || textBox15.Text == ","
                    || textBox16.Text == ",")
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
            else if (dgridzlyformat == true)
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



                dataGridView11.RowCount = nrlini;
                dataGridView22.RowCount = nrlini;
                dataGridView33.RowCount = nrlini;
                dataGridView44.RowCount = nrlini + 1;
                dataGridView1.RowCount = 4;

                double[] dlodc = new double[nrlini];
                double[] przkabla = new double[nrlini];
                double[] reaktjedn = new double[nrlini];
                double[] moc = new double[nrlini];
                double[] wspmoc = new double[nrlini];
                string[] chtrobc = new string[nrlini];


                double[] sinus = new double[nrlini];
                double[] cosinus = new double[nrlini];

                Complex[] pradodb = new Complex[nrlini];
                Complex[] impedancjalini = new Complex[nrlini];
                Complex impedancjatrafo1; //poczatkowy T1
                Complex impedancjatrafo2; // końcowy T2

                double przewodnoscwlasc = Convert.ToDouble(textBox6.Text);
                double napznam = Convert.ToDouble(textBox3.Text) * 1000;
                double przewodnoscwlasc1 = Convert.ToDouble(textBox6.Text);
                double napstart = Convert.ToDouble(textBox1.Text) * 1000;
                double napend = Convert.ToDouble(textBox4.Text) * 1000;
                double kontstart = Convert.ToDouble(textBox2.Text);
                double kontend = Convert.ToDouble(textBox5.Text);
                Complex napzesp1 = new Complex(0, 0);
                Complex napzesp2 = new Complex(0, 0);
                Complex napzesp1znam = new Complex(0, 0);
                Complex napzesp2znam = new Complex(0, 0);
                Complex roznicanapiec12 = new Complex(0, 0);
                Complex roznicanapiec21 = new Complex(0, 0);
                Complex pradzasilajacy1 = new Complex(0, 0);
                Complex pradzasilajacy2 = new Complex(0, 0);

                Complex[] spadeknap = new Complex[nrlini];
                Double[] napwpkt = new double[nrlini + 1];
                //fdfddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddd

                double kontstartcos = Math.Cos(kontstart * Math.PI / 180);
                double kontstartsin = Math.Sin(kontstart * Math.PI / 180);
                double kontendcos = Math.Cos(kontend * Math.PI / 180);
                double kontendsin = Math.Sin(kontend * Math.PI / 180);



                napzesp1.real = napstart * kontstartcos;
                napzesp1.imaginary = napstart * kontstartsin;

                napzesp2.real = napend * kontendcos;
                napzesp2.imaginary = napend * kontendsin;


                for (int h = 1; h < nrlini; h++)
                {
                    dataGridView11.Rows[h - 1].HeaderCell.Value = "I" + (h);
                }

                for (int h = 0; h < nrlini + 1; h++)
                {
                    dataGridView44.Rows[h].HeaderCell.Value = "U" + h;
                }


                dataGridView11.ReadOnly = true;
                dataGridView22.ReadOnly = true;
                dataGridView33.ReadOnly = true;
                dataGridView44.ReadOnly = true;
                dataGridView1.ReadOnly = true;

                dataGridView1.AllowUserToResizeRows = false;

                for (int i = 0; i < nrlini; i++)
                {
                    dlodc[i] = Convert.ToDouble(dGrid1[1, i].Value);
                    przkabla[i] = Convert.ToDouble(dGrid1[2, i].Value);
                    reaktjedn[i] = Convert.ToDouble(dGrid1[3, i].Value);

                }

                for (int h = 1; h < nrlini; h++)
                {
                    if (checkBox1.Checked == true && h == nrlini - 1)
                    {
                        moc[h] = Convert.ToDouble(dGrid1[5, h].Value) * 1000;
                        wspmoc[h] = Convert.ToDouble(dGrid1[6, h].Value);
                        chtrobc[h] = (string)dGrid1[7, h].Value;
                    }
                    else
                    {
                        moc[h] = Convert.ToDouble(dGrid1[5, h].Value) * 1000;
                        wspmoc[h] = Convert.ToDouble(dGrid1[6, h].Value);
                        chtrobc[h] = (string)dGrid1[7, h].Value;

                    }

                    cosinus[h] = wspmoc[h];
                    sinus[h] = Math.Sqrt(1 - Math.Pow(cosinus[h], 2));
                    pradodb[h].real = (moc[h] / (Math.Sqrt(3) * napznam * cosinus[h])) * cosinus[h];

                    if (chtrobc[h] == "Indukcyjny")
                    {
                        pradodb[h].imaginary = (moc[h] / (Math.Sqrt(3) * napznam * cosinus[h])) * (sinus[h] * (-1));
                    }
                    else
                    {
                        pradodb[h].imaginary = (moc[h] / (Math.Sqrt(3) * napznam * cosinus[h])) * sinus[h];
                    }

                    pradodb[h].real = Math.Round(pradodb[h].real, 2);
                    pradodb[h].imaginary = Math.Round(pradodb[h].imaginary, 2);
                    dataGridView11.RowCount = nrlini;
                    dataGridView11[0, h - 1].Value = pradodb[h].ToString();



                }

                //double rezyst;
                //double reakt;
                for (int l = 0; l < nrlini; l++)
                {
                    impedancjalini[l].real = (1000 / (przewodnoscwlasc * przkabla[l])) * dlodc[l];
                    impedancjalini[l].imaginary = reaktjedn[l] * dlodc[l];
                    if (checkBox1.Checked)
                    {

                        double napgora1 = Convert.ToDouble(textBox10.Text);
                        double napdol1 = Convert.ToDouble(textBox11.Text);
                        double napgora2 = Convert.ToDouble(textBox15.Text);
                        double napdol2 = Convert.ToDouble(textBox16.Text);
                        double deltapcutrafo1 = Convert.ToDouble(textBox7.Text);
                        double napzwarciatrafo1 = Convert.ToDouble(textBox8.Text);
                        double moctrafo1 = Convert.ToDouble(textBox9.Text);
                        //napięcie na znamionowe
                        napzesp1znam = napzesp1 * napdol1 / napgora1;
                        napzesp2znam = napzesp2 * napdol2 / napgora2;



                        impedancjatrafo1.real = (deltapcutrafo1 * Math.Pow(napdol1, 2)) / (100 * moctrafo1);
                        impedancjatrafo1.imaginary = (napzwarciatrafo1 * Math.Pow(napdol1, 2)) / (100 * moctrafo1);

                        impedancjalini[0].real = impedancjatrafo1.real;
                        impedancjalini[0].imaginary = impedancjatrafo1.imaginary;


                        double deltapcutrafo2 = Convert.ToDouble(textBox12.Text);
                        double napzwarciatrafo2 = Convert.ToDouble(textBox13.Text);
                        double moctrafo2 = Convert.ToDouble(textBox14.Text);
                        impedancjatrafo2.real = (deltapcutrafo2 * Math.Pow(napdol2, 2)) / (100 * moctrafo2);
                        impedancjatrafo2.imaginary = (napzwarciatrafo2 * Math.Pow(napdol2, 2)) / (100 * moctrafo2);

                        impedancjalini[nrlini - 1].real = impedancjatrafo2.real;
                        impedancjalini[nrlini - 1].imaginary = impedancjatrafo2.imaginary;

                    }


                }
                Complex impedancjacalkowita = new Complex(0, 0);
                Complex pradimpedancja12 = new Complex(0, 0);
                Complex pradimpedancja21 = new Complex(0, 0);
                Complex pradzasilajacygora1 = new Complex(0, 0);
                Complex pradzasilajacygora2 = new Complex(0, 0);

                for (int i = 0; i < nrlini; i++)
                {
                    impedancjacalkowita += impedancjalini[i];


                    if (i + 1 < nrlini)
                    {
                        pradimpedancja21 += pradodb[i + 1] * (impedancjacalkowita);
                        
                    }



                }
                Complex imptymczas = impedancjacalkowita;
                int pktsplywu = 0;
                for (int i = 1; i < nrlini; i++)
                {



                    imptymczas -= impedancjalini[i - 1];
                    pradimpedancja12 += pradodb[i] * (imptymczas);
                    

                }






                Complex pradwyr1 = new Complex(0, 0);
                Complex pradwyr2 = new Complex(0, 0);
                if (checkBox1.Checked)
                {
                    double napgora1 = Convert.ToDouble(textBox10.Text);
                    double napdol1 = Convert.ToDouble(textBox11.Text);
                    double napgora2 = Convert.ToDouble(textBox15.Text);
                    double napdol2 = Convert.ToDouble(textBox16.Text);
                    roznicanapiec12 = (napzesp1znam - napzesp2znam);
                    roznicanapiec21 = (napzesp2znam - napzesp1znam);
                    pradwyr1 = roznicanapiec12 / (impedancjacalkowita * Math.Sqrt(3));
                    pradwyr2 = roznicanapiec21 / (impedancjacalkowita * Math.Sqrt(3));
                    pradzasilajacy1 = (pradimpedancja12 / impedancjacalkowita);
                    pradzasilajacy2 = (pradimpedancja21 / impedancjacalkowita);

                    pradzasilajacy1.real = Math.Round(pradzasilajacy1.real, 2);
                    pradzasilajacy1.imaginary = Math.Round(pradzasilajacy1.imaginary, 2);

                    pradzasilajacy2.real = Math.Round(pradzasilajacy2.real, 2);
                    pradzasilajacy2.imaginary = Math.Round(pradzasilajacy2.imaginary, 2);
                    napwpkt[0] = napzesp1znam.real;
                    napwpkt[nrlini] = napzesp2znam.real;
                    pradzasilajacygora1 = (pradzasilajacy1 + pradwyr1) * (napdol1 / napgora1);
                    pradzasilajacygora2 = (pradzasilajacy2 + pradwyr2) * (napdol2 / napgora2);
                    pradzasilajacygora1.real = Math.Round(pradzasilajacygora1.real, 2);
                    pradzasilajacygora1.imaginary = Math.Round(pradzasilajacygora1.imaginary, 2);
                    pradzasilajacygora2.real = Math.Round(pradzasilajacygora2.real, 2);
                    pradzasilajacygora2.imaginary = Math.Round(pradzasilajacygora2.imaginary, 2);
                    dataGridView1.RowHeadersWidth = 70;
                    dataGridView1.Visible = true;
                    Complex pradzaswyr1 = new Complex(0, 0);
                    Complex pradzaswyr2 = new Complex(0, 0);
                    pradzaswyr1 = pradzasilajacy1 + pradwyr1;
                    pradzaswyr2 = pradzasilajacy2 + pradwyr2;
                    pradzaswyr1.real = Math.Round(pradzaswyr1.real, 2);
                    pradzaswyr1.imaginary = Math.Round(pradzaswyr1.imaginary, 2);
                    pradzaswyr2.real = Math.Round(pradzaswyr2.real, 2);
                    pradzaswyr2.imaginary = Math.Round(pradzaswyr2.imaginary, 2);
                    dataGridView1[0, 0].Value = pradzaswyr1;
                    dataGridView1[0, 1].Value = pradzaswyr2;
                    dataGridView1[0, 2].Value = pradzasilajacygora1;
                    dataGridView1[0, 3].Value = pradzasilajacygora2;
                    dataGridView1.Rows[0].HeaderCell.Value = "Iz0" + "(" + napdol1 + ")";
                    dataGridView1.Rows[1].HeaderCell.Value = "Iz" + nrlini + "(" + napdol2 + ")";
                    dataGridView1.Rows[2].HeaderCell.Value = "Iz0" + "(" + napgora1 + ")";
                    dataGridView1.Rows[3].HeaderCell.Value = "Iz" + nrlini + "(" + napgora2 + ")";
                }
                else if (!checkBox1.Checked)
                {
                    roznicanapiec12 = (napzesp1 - napzesp2);
                    roznicanapiec21 = (napzesp2 - napzesp1);
                    pradwyr1 = roznicanapiec12 / (impedancjacalkowita * Math.Sqrt(3));
                    pradwyr2 = roznicanapiec21 / (impedancjacalkowita * Math.Sqrt(3));
                   
                    pradzasilajacy1 = (pradimpedancja12 / impedancjacalkowita);
                    pradzasilajacy2 = (pradimpedancja21 / impedancjacalkowita);
                   
                    pradzasilajacy1.real = Math.Round(pradzasilajacy1.real, 2);
                    pradzasilajacy1.imaginary = Math.Round(pradzasilajacy1.imaginary, 2);

                    pradzasilajacy2.real = Math.Round(pradzasilajacy2.real, 2);
                    pradzasilajacy2.imaginary = Math.Round(pradzasilajacy2.imaginary, 2);
                    napwpkt[0] = napzesp1.real;
                    napwpkt[nrlini] = napzesp2.real;
                    pradzasilajacygora1 = pradzasilajacy1;
                    pradzasilajacygora2 = pradzasilajacy2;
                    pradzasilajacygora1.real = Math.Round(pradzasilajacygora1.real, 2);
                    pradzasilajacygora1.imaginary = Math.Round(pradzasilajacygora1.imaginary, 2);
                    pradzasilajacygora2.real = Math.Round(pradzasilajacygora2.real, 2);
                    pradzasilajacygora2.imaginary = Math.Round(pradzasilajacygora2.imaginary, 2);
                    dataGridView1[0, 0].Value = pradzasilajacy1;
                    dataGridView1[0, 1].Value = pradzasilajacy2;
                    dataGridView1[0, 2].Value = pradzasilajacygora1;
                    dataGridView1[0, 3].Value = pradzasilajacygora2;



                }
                Complex[] rozplywprad12 = new Complex[nrlini];
                Complex[] rozplywprad21 = new Complex[nrlini];
                if (roznicanapiec12.real >= 0)
                {
                    for (int t = 0; t < nrlini; t++)
                    {
                        if (t == 0)
                        {
                            rozplywprad12[t] = pradzasilajacy1 + pradwyr1;
                            


                        }
                        else
                        {
                            
                            rozplywprad12[t] = rozplywprad12[t - 1] - pradodb[t];
                            

                        }
                    }
                    for (int z = 0; z < nrlini; z++)
                    {
                        if (rozplywprad12[z].real <= 0)
                        {
                            pktsplywu = z;
                            break;
                        }
                        else
                        {
                            if (roznicanapiec12.real > 0)
                            {
                                pktsplywu = nrlini;
                            }
                            else
                            {
                                pktsplywu = 0;
                            }
                        }
                    }
                    for (int z = 0; z < nrlini; z++)
                    {
                        if (rozplywprad12[z].real < 0)
                        {
                            rozplywprad21[z] = rozplywprad12[z] * (-1);
                        }
                    }
                }
                else
                {
                    for (int t = nrlini - 1; t >= 0; t--)
                    {
                        if (t == nrlini - 1)
                        {
                            rozplywprad12[t] = pradzasilajacy2 + pradwyr2;
                            
                        }
                        else
                        {
                            
                            rozplywprad12[t] = rozplywprad12[t + 1] - pradodb[t + 1];
                            
                        }



                    }
                    for (int z = nrlini - 1; z >= 0; z--)
                    {
                        if (rozplywprad12[z].real <= 0)
                        {
                            pktsplywu = z;
                            break;
                        }
                        else
                        {
                            if (roznicanapiec12.real > 0)
                            {
                                pktsplywu = nrlini;

                            }
                            else
                            {
                                pktsplywu = 0;

                            }
                        }
                    }
                    for (int z = 0; z < nrlini; z++)
                    {
                        if (rozplywprad12[z].real < 0)
                        {
                            rozplywprad21[z] = rozplywprad12[z] * (-1);
                        }
                    }
                }




                
                label36.Visible = true;
                label36.Text = "Punkt spływu: " + pktsplywu;

                dataGridView22.RowCount = nrlini;
                dataGridView33.RowCount = nrlini;
                for (int t = 0; t < nrlini; t++)
                {
                    rozplywprad12[t].real = Math.Round(rozplywprad12[t].real, 2);
                    rozplywprad12[t].imaginary = Math.Round(rozplywprad12[t].imaginary, 2);
                    rozplywprad21[t].real = Math.Round(rozplywprad21[t].real, 2);
                    rozplywprad21[t].imaginary = Math.Round(rozplywprad21[t].imaginary, 2);
                    dataGridView22[0, t].Value = rozplywprad12[t].ToString();
                    if (rozplywprad12[t].real < 0)
                    {
                        dataGridView22[0, t].Value = rozplywprad21[t].ToString();
                    }

                }

                Complex[] spadeknap2 = new Complex[nrlini];
                for (int s = 0; s < nrlini; s++)
                {

                    spadeknap[s] = rozplywprad12[s] * impedancjalini[s];
                    spadeknap[s].real = spadeknap[s].real * Math.Sqrt(3);
                    spadeknap[s].imaginary = spadeknap[s].imaginary * Math.Sqrt(3);
                    
                    spadeknap[s].real = Math.Round(spadeknap[s].real, 2);
                    spadeknap[s].imaginary = Math.Round(spadeknap[s].imaginary, 2);
                    if (spadeknap[s].real < 0)
                    {
                        spadeknap2[s] = spadeknap[s] * (-1);
                        dataGridView33[0, s].Value = spadeknap2[s].real.ToString();
                    }
                    else
                    {
                        dataGridView33[0, s].Value = spadeknap[s].real.ToString();
                    }

                }

                for (int d = 1; d < nrlini; d++)
                {



                    if (roznicanapiec12.real < 0)
                    {
                        for (int t = nrlini - 1; t > 0; t--)
                        {
                            napwpkt[t] = napwpkt[t + 1] - spadeknap[t].real;
                        }
                    }
                    else
                    {
                        napwpkt[d] = napwpkt[d - 1] - spadeknap[d - 1].real;
                    }

                    


                }
                for (int d = 0; d < nrlini + 1; d++)
                {
                    napwpkt[d] = Math.Round(napwpkt[d], 2);

                    dataGridView44[0, d].Value = napwpkt[d].ToString();
                }
                for (int h = 0; h < nrlini; h++)
                {

                    if (h >= pktsplywu)
                    {
                        dataGridView22.Rows[h].HeaderCell.Value = "I" + (h + 1) + h;
                        dataGridView33.Rows[h].HeaderCell.Value = "ΔU" + (h + 1) + h;
                    }
                    else
                    {
                        dataGridView22.Rows[h].HeaderCell.Value = "I" + h + (h + 1);
                        dataGridView33.Rows[h].HeaderCell.Value = "ΔU" + h + (h + 1);
                    }









                }

            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            DataGridViewComboBoxColumn column = new DataGridViewComboBoxColumn();
            column.HeaderText = "Ch-ter obciążenia";
            column.Name = " ";
            column.Items.AddRange("Indukcyjny", "Pojemnościowy");
            numericUpDown1.Value = 2;
            for (int i = 1; i < 4; i++)
            {
                dGrid1[i, 0].Value = null;
                dGrid1[i, 1].Value = null;
            }
            for (int i = 5; i < 7; i++)
            {
                dGrid1[i, 1].Value = null;
            }
            dGrid1.Columns.RemoveAt(7);
            dGrid1.Columns.Add(column);
            dGrid1[7, 0].ReadOnly = true;
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;
            textBox8.Text = null;
            textBox9.Text = null;
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;
            textBox14.Text = null;
            textBox15.Text = null;
            textBox16.Text = null;

            checkBox1.Checked = false;

            dataGridView11.RowCount = 1;
            dataGridView11[0, 0].Value = null;
            dataGridView22.RowCount = 1;
            dataGridView22[0, 0].Value = null;
            dataGridView33.RowCount = 1;
            dataGridView33[0, 0].Value = null;
            dataGridView44.RowCount = 1;
            dataGridView44[0, 0].Value = null;
            pictureBox1.Image = null;
            dataGridView11.Rows[0].HeaderCell.Value = "";
            dataGridView22.Rows[0].HeaderCell.Value = "";
            dataGridView33.Rows[0].HeaderCell.Value = "";
            dataGridView44.Rows[0].HeaderCell.Value = "";
            dataGridView1.RowCount = 1;
            dataGridView1[0, 0].Value = null;
            dataGridView1.Rows[0].HeaderCell.Value = "";
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label32_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label35_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void button3_Click(object sender, EventArgs e)
        {
            int iloscodb = (int)numericUpDown1.Value - 1;

            string odbnr;
            string path = Directory.GetCurrentDirectory();

            string fullpath = path + "\\" + "obrazki\\";
            if (!checkBox1.Checked)
            {
                odbnr = "d" + iloscodb;
                Image image = Image.FromFile(fullpath + odbnr + ".jpg");
                this.pictureBox1.Image = image;
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            }
            else if (checkBox1.Checked)
            {
                odbnr = "d" + iloscodb + "t";
                Image image = Image.FromFile(fullpath + odbnr + ".jpg");
                this.pictureBox1.Image = image;
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image == null)
            {
                string message = "Brak układu, załaduj układ";
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

        private void dGrid1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            /*int nrlini = (int)numericUpDown1.Value;
            if (!checkBox1.Checked)
            {
                for (int kolumna = 1; kolumna < 4; kolumna++)
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
                for (int kolumna = 5; kolumna < 7; kolumna++)
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
            }
            else if (checkBox1.Checked)
            {
                for (int kolumna = 1; kolumna < 4; kolumna++)
                {
                    for (int wiersz = 1; wiersz < nrlini-1; wiersz++)
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

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox14_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox15_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox16_KeyPress(object sender, KeyPressEventArgs e)
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
            string fullpath = path + "\\" + "przyklady\\przyklad1d.txt";
            string text = File.ReadAllText(fullpath);


            string[] words = text.Split(new char[] { ' ', '\n' });
            textBox1.Text = Convert.ToString(words[0]);
            textBox4.Text = Convert.ToString(words[1]);
            textBox3.Text = Convert.ToString(words[2]);
            textBox2.Text = Convert.ToString(words[3]);
            textBox5.Text = Convert.ToString(words[4]);
            textBox6.Text = Convert.ToString(words[5]);

            string wspmoc1 = "Indukcyjny";
            string wspmoc2 = "Indukcyjny";
            string wspmoc3 = "Indukcyjny";
            string[] odc1 = { "1", "0,5", "95", "0,36" };
            int nrlini = Convert.ToInt32(words[6]);
            numericUpDown1.Value = nrlini;

            for (int kolumna = 0; kolumna < 7; kolumna++)
            {
                for (int wiersz = 1; wiersz < nrlini; wiersz++)
                {
                    dGrid1[kolumna, wiersz].Value = words[0 + wiersz * 7 + kolumna];


                }
            }
            dGrid1[7, 1].Value = wspmoc1;
            dGrid1[7, 2].Value = wspmoc2;
            dGrid1[7, 3].Value = wspmoc3;
            dGrid1[0, 0].Value = odc1[0];
            dGrid1[1, 0].Value = odc1[1];
            dGrid1[2, 0].Value = odc1[2];
            dGrid1[3, 0].Value = odc1[3];
        }

        private void button_przyklad2_Click(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory();
            string fullpath = path + "\\" + "przyklady\\przyklad2d.txt";
            string text1 = File.ReadAllText(fullpath);


            string[] words1 = text1.Split(new char[] { ' ', '\n' });
            textBox1.Text = Convert.ToString(words1[0]);
            textBox4.Text = Convert.ToString(words1[1]);
            textBox3.Text = Convert.ToString(words1[2]);
            textBox2.Text = Convert.ToString(words1[3]);
            textBox5.Text = Convert.ToString(words1[4]);
            textBox6.Text = Convert.ToString(words1[5]);
            string wspmoc1 = "Indukcyjny";
            string wspmoc2 = "Indukcyjny";
            string wspmoc3 = "Indukcyjny";
            int nrlini = Convert.ToInt32(words1[6]);
            numericUpDown1.Value = nrlini;
            for (int kolumna = 0; kolumna < 7; kolumna++)
            {
                for (int wiersz = 1; wiersz < nrlini; wiersz++)
                {
                    dGrid1[kolumna, wiersz].Value = words1[10 + wiersz * 7 + kolumna];


                }
            }
            dGrid1[7, 1].Value = wspmoc1;
            dGrid1[7, 2].Value = wspmoc2;
            dGrid1[7, 3].Value = wspmoc3;
            checkBox1.Checked = true;
            textBox7.Text = Convert.ToString(words1[7]);
            textBox8.Text = Convert.ToString(words1[8]);
            textBox9.Text = Convert.ToString(words1[9]);
            textBox10.Text = Convert.ToString(words1[10]);
            textBox11.Text = Convert.ToString(words1[11]);

            textBox12.Text = Convert.ToString(words1[12]);
            textBox13.Text = Convert.ToString(words1[13]);
            textBox14.Text = Convert.ToString(words1[14]);
            textBox15.Text = Convert.ToString(words1[15]);
            textBox16.Text = Convert.ToString(words1[16]);
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }
        private void copyAlltoClipboard1()
        {
            dataGridView11.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView11.MultiSelect = true;
            dataGridView11.SelectAll();
            DataObject dataObj1 = dataGridView11.GetClipboardContent();
            if (dataObj1 != null)
                Clipboard.SetDataObject(dataObj1);

        }
        private void copyAlltoClipboard2()
        {
            dataGridView22.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView22.MultiSelect = true;
            dataGridView22.SelectAll();
            DataObject dataObj2 = dataGridView22.GetClipboardContent();
            if (dataObj2 != null)
                Clipboard.SetDataObject(dataObj2);

        }
        private void copyAlltoClipboard3()
        {

            dataGridView33.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView33.MultiSelect = true;
            dataGridView33.SelectAll();
            DataObject dataObj3 = dataGridView33.GetClipboardContent();
            if (dataObj3 != null)
                Clipboard.SetDataObject(dataObj3);

        }
        private void copyAlltoClipboard4()
        {
            dataGridView44.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView44.MultiSelect = true;
            dataGridView44.SelectAll();
            DataObject dataObj4 = dataGridView44.GetClipboardContent();
            if (dataObj4 != null)
                Clipboard.SetDataObject(dataObj4);

        }
        private void copyAlltoClipboard5()
        {
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView1.MultiSelect = true;
            dataGridView1.SelectAll();
            DataObject dataObj5 = dataGridView1.GetClipboardContent();
            if (dataObj5 != null)
                Clipboard.SetDataObject(dataObj5);
        }
        private void button5_Click(object sender, EventArgs e)
        {


            if (dataGridView1.Rows[0].Cells[0].Value == null)
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
    }
}
