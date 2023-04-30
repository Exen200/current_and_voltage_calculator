using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Numerics;
using System.Drawing.Imaging;
using Excel = Microsoft.Office.Interop.Excel;

namespace WinFormsApp1
{
    public partial class UserControl3 : UserControl
    {
        public UserControl3()
        {
            InitializeComponent();
           
        }

        private void UserControl3_Load(object sender, EventArgs e)
        {
            
            // pierwszy dGrid1 - linia zasilajaca
            dGrid1.RowCount = 1; // nadanie ilości wierszy w datagrid
            dGrid1.ColumnCount = 4; // nadanie ilości kolumn w datagrid
            dGrid1.Columns[0].HeaderCell.Value = "Nr odcinka"; // przypisanie nazwy do nagłówka danej kolumny
            dGrid1.Columns[1].HeaderCell.Value = "Dł.odcinka [km]"; // -||-
            dGrid1.Columns[2].HeaderCell.Value = "Prz. kabla/linii [mm^2]";
            dGrid1.Columns[3].HeaderCell.Value = "X'[Ω/km]";
            dGrid1[0, 0].Value = 1;
            dGrid1[0, 0].ReadOnly = true;
            dGrid1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;  // blokowanie sortowania, żeby po kliknięciu w nagłówek nie zmieniało
            dGrid1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;  // się ustawienie danych tabeli, 
            dGrid1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;  // np. z 1 2 3 4 na 4 3 2 1, widać to po numerach
            dGrid1.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;

            // drugi dGrid2 - gałąź 1
            dGrid2.RowCount = 1; 
            dGrid2.ColumnCount = 7; 
            dGrid2.Columns[0].HeaderCell.Value = "Nr odcinka"; 
            dGrid2.Columns[1].HeaderCell.Value = "Dł.odcinka [km]"; 
            dGrid2.Columns[2].HeaderCell.Value = "Prz. kabla/linii [mm^2]";
            dGrid2.Columns[3].HeaderCell.Value = "X'[Ω/km]";
            dGrid2.Columns[4].HeaderCell.Value = "Nr odbioru";
            dGrid2.Columns[5].HeaderCell.Value = "Moc [kW]";
            dGrid2.Columns[6].HeaderCell.Value = "Wsp. mocy [-]";
            
            DataGridViewComboBoxColumn columng1 = new DataGridViewComboBoxColumn(); // stworzenie nowej kolumny o właściwościach comboboxa
            columng1.HeaderText = "Ch-ter obciążenia"; // przypisanie nazwy do nagłówka nowej kolumny
            columng1.Name = " "; // przypisanie nazwy nowej kolumnie
            columng1.Items.AddRange("Indukcyjny", "Pojemnościowy"); // dodanie wartości do nowej kolumny typu comboBox
            dGrid2.Columns.Add(columng1); // dodanie tej kolumny do datagrid

            dGrid2[0, 0].Value = 1;
            dGrid2[4, 0].Value = 1;
            dGrid2[0, 0].ReadOnly = true;
            dGrid2[4, 0].ReadOnly = true;
            dataGridView4.ReadOnly = true;

            dGrid2.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;  
            dGrid2.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;  
            dGrid2.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;  
            dGrid2.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid2.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid2.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid2.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid2.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;

            // trzeci dGrid3 - gałąź 2
            dGrid3.RowCount = 1;
            dGrid3.ColumnCount = 7;
            dGrid3.Columns[0].HeaderCell.Value = "Nr odcinka";
            dGrid3.Columns[1].HeaderCell.Value = "Dł.odcinka [km]";
            dGrid3.Columns[2].HeaderCell.Value = "Prz. kabla/linii [mm^2]";
            dGrid3.Columns[3].HeaderCell.Value = "X'[Ω/km]";
            dGrid3.Columns[4].HeaderCell.Value = "Nr odbioru";
            dGrid3.Columns[5].HeaderCell.Value = "Moc [kW]";
            dGrid3.Columns[6].HeaderCell.Value = "Wsp. mocy [-]";
           
            DataGridViewComboBoxColumn columng2 = new DataGridViewComboBoxColumn(); 
            columng2.HeaderText = "Ch-ter obciążenia"; 
            columng2.Name = " "; 
            columng2.Items.AddRange("Indukcyjny", "Pojemnościowy"); 
            dGrid3.Columns.Add(columng2); 

            dGrid3[0, 0].Value = 1;
            dGrid3[4, 0].Value = 1;
            dGrid3[0, 0].ReadOnly = true;
            dGrid3[4, 0].ReadOnly = true;

            dGrid3.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid3.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid3.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid3.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid3.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid3.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid3.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid3.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;

            // czwarty dGrid4 - gałąź 3
            dGrid4.RowCount = 1;
            dGrid4.ColumnCount = 7;
            dGrid4.Columns[0].HeaderCell.Value = "Nr odcinka";
            dGrid4.Columns[1].HeaderCell.Value = "Dł.odcinka [km]";
            dGrid4.Columns[2].HeaderCell.Value = "Prz. kabla/linii [mm^2]";
            dGrid4.Columns[3].HeaderCell.Value = "X'[Ω/km]";
            dGrid4.Columns[4].HeaderCell.Value = "Nr odbioru";
            dGrid4.Columns[5].HeaderCell.Value = "Moc [kW]";
            dGrid4.Columns[6].HeaderCell.Value = "Wsp. mocy [-]";

            DataGridViewComboBoxColumn columng3 = new DataGridViewComboBoxColumn();
            columng3.HeaderText = "Ch-ter obciążenia";
            columng3.Name = " ";
            columng3.Items.AddRange("Indukcyjny", "Pojemnościowy");
            dGrid4.Columns.Add(columng3);

            dGrid4[0, 0].Value = 1;
            dGrid4[4, 0].Value = 1;
            dGrid4[0, 0].ReadOnly = true;
            dGrid4[4, 0].ReadOnly = true;

            dGrid4.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid4.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid4.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid4.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid4.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid4.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid4.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
            dGrid4.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;

            dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView2.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView3.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView12.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView4.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            int nrlinig1 = (int)numericUpDown1.Value;  //ilość odcinków
            dGrid2.RowCount = nrlinig1;         // zmiana ilości odcinków

            for (int i = 0; i < nrlinig1; i++) //pętla do tworzenia numeracji w niżej podanych komórkach
            {
                dGrid2[0, i].Value = i + 1; //wartość w komórce [0,i] jest równe i + 1 i po każdej pętli zmnienia się
                dGrid2[4, i].Value = i + 1; // -||-, komórka oraz jej wartość

                dGrid2[0, i].ReadOnly = true;  //komórki z numeracją są readonly, więc nie można ich edytować
                dGrid2[4, i].ReadOnly = true;

            }
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            int nrlinig2 = (int)numericUpDown2.Value;  //ilość odcinków
            dGrid3.RowCount = nrlinig2;         // zmiana ilości odcinków

            for (int i = 0; i < nrlinig2; i++) //pętla do tworzenia numeracji w niżej podanych komórkach
            {
                dGrid3[0, i].Value = i + 1; //wartość w komórce [0,i] jest równe i + 1 i po każdej pętli zmnienia się
                dGrid3[4, i].Value = i + 1; // -||-, komórka oraz jej wartość

                dGrid3[0, i].ReadOnly = true;  //komórki z numeracją są readonly, więc nie można ich edytować
                dGrid3[4, i].ReadOnly = true;

            }
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            int nrlinig3 = (int)numericUpDown3.Value;  //ilość odcinków
            dGrid4.RowCount = nrlinig3;         // zmiana ilości odcinków

            for (int i = 0; i < nrlinig3; i++) //pętla do tworzenia numeracji w niżej podanych komórkach
            {
                dGrid4[0, i].Value = i + 1; //wartość w komórce [0,i] jest równe i + 1 i po każdej pętli zmnienia się
                dGrid4[4, i].Value = i + 1; // -||-, komórka oraz jej wartość

                dGrid4[0, i].ReadOnly = true;  //komórki z numeracją są readonly, więc nie można ich edytować
                dGrid4[4, i].ReadOnly = true;

            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) //jezeli checkbox zostanie klikniety
            {
                label11.Visible = true; // pojawia sie label - wprowadz dane transformatora
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

                textBox4.Visible = true;
                textBox5.Visible = true;
                textBox6.Visible = true;
                textBox7.Visible = true;
                textBox8.Visible = true;
                // linia zasilająca = transformator
                dGrid1[1, 0].ReadOnly = true; // długość odcinka w tej komórce -> readonly 
                dGrid1[2, 0].ReadOnly = true;  // przekroj odcinka w tej komórce -> readonly
                dGrid1[3, 0].ReadOnly = true;  // reaktancjajedn odcinka w tej komórce -> readonly

                dGrid1[1, 0].Value = null; // wartość tych komórek jest zmieniana na null, czyli brak wartości
                dGrid1[2, 0].Value = null;
                dGrid1[3, 0].Value = null;
            }
            else
            {
                label11.Visible = false; // pojawia sie label - wprowadz dane transformatora
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
                //linia zasilajaca != transformator
                dGrid1[1, 0].ReadOnly = false;
                dGrid1[2, 0].ReadOnly = false;
                dGrid1[3, 0].ReadOnly = false;

                dGrid1[1, 0].Value = null; // wartość tych komórek jest zmieniana na null, czyli brak wartości
                dGrid1[2, 0].Value = null;
                dGrid1[3, 0].Value = null;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.RowHeadersWidth = 70;
            dataGridView2.RowHeadersWidth = 70;
            dataGridView3.RowHeadersWidth = 70;
            dataGridView12.RowHeadersWidth = 70;
            dataGridView4.RowHeadersWidth = 100;
            if (!checkBox1.Checked)
            {
                dataGridView4.Visible = false;
            }
            int nrlinig1 = (int)numericUpDown1.Value;
            int nrlinig2 = (int)numericUpDown2.Value;
            int nrlinig3 = (int)numericUpDown3.Value;

            bool brakkolumnzas = false;
            bool brakkolumn17g1 = false;
            bool brakkolumn17g2 = false;
            bool brakkolumn17g3 = false;
            bool braktxtboxa = false;
            bool txtboxzlyformat = false;
            bool dgridzlyformat = false;
            bool ujemnelubzero = false;
            
            if (!checkBox1.Checked)
            {   
                for (int kolumna = 0; kolumna < 4; kolumna++)
                {
                    for (int wiersz = 0; wiersz < 1; wiersz++)
                    {
                        if (dGrid1[kolumna, wiersz].Value == null)
                        {
                            brakkolumnzas = true;
                        }
                    }
                }
                for (int kolumna = 0; kolumna < 4; kolumna++)
                {
                    for (int wiersz = 0; wiersz < 1; wiersz++)
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
                    for (int wiersz = 0; wiersz < 1; wiersz++)
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


                for (int kolumna = 0; kolumna < 8; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig1; wiersz++)
                    {
                        if (dGrid2[kolumna, wiersz].Value == null)
                        {
                            brakkolumn17g1 = true;
                        }
                    }
                }
                for (int kolumna = 0; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig1; wiersz++)
                    {
                        if (dGrid2[kolumna, wiersz].Value != null)
                        {
                            string? value = dGrid2.Rows[wiersz].Cells[kolumna].Value.ToString();
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
                    for (int wiersz = 0; wiersz < nrlinig1; wiersz++)
                    {
                        if (dgridzlyformat == false)
                        {
                            double a;
                            a = Convert.ToDouble(dGrid2[kolumna, wiersz].Value);

                            if (a <= 0)
                            {
                                ujemnelubzero = true;
                            }
                        }
                    }
                }


                for (int kolumna = 0; kolumna < 8; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig2; wiersz++)
                    {
                        if (dGrid3[kolumna, wiersz].Value == null)
                        {
                            brakkolumn17g2 = true;
                        }
                    }
                }
                for (int kolumna = 0; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig2; wiersz++)
                    {
                        if (dGrid3[kolumna, wiersz].Value != null)
                        {
                            string? value = dGrid3.Rows[wiersz].Cells[kolumna].Value.ToString();
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
                    for (int wiersz = 0; wiersz < nrlinig2; wiersz++)
                    {
                        if (dgridzlyformat == false)
                        {
                            double a;
                            a = Convert.ToDouble(dGrid3[kolumna, wiersz].Value);

                            if (a <= 0)
                            {
                                ujemnelubzero = true;
                            }
                        }
                    }
                }


                for (int kolumna = 0; kolumna < 8; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig3; wiersz++)
                    {
                        if (dGrid4[kolumna, wiersz].Value == null)
                        {
                            brakkolumn17g3 = true;
                        }
                    }
                }
                for (int kolumna = 0; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig3; wiersz++)
                    {
                        if (dGrid4[kolumna, wiersz].Value != null)
                        {
                            string? value = dGrid4.Rows[wiersz].Cells[kolumna].Value.ToString();
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
                    for (int wiersz = 0; wiersz < nrlinig3; wiersz++)
                    {
                        if (dgridzlyformat == false)
                        {
                            double a;
                            a = Convert.ToDouble(dGrid4[kolumna, wiersz].Value);

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
                if (textBox1.Text == "0," || textBox2.Text == "0," || textBox3.Text == "0,")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == ",0" || textBox2.Text == ",0" || textBox3.Text == ",0")
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
                for (int kolumna = 0; kolumna < 8; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig1; wiersz++)
                    {
                        if (dGrid2[kolumna, wiersz].Value == null)
                        {
                            brakkolumn17g1 = true;
                        }
                    }
                }
                for (int kolumna = 0; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig1; wiersz++)
                    {
                        if (dGrid2[kolumna, wiersz].Value != null)
                        {
                            string? value = dGrid2.Rows[wiersz].Cells[kolumna].Value.ToString();
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
                    for (int wiersz = 0; wiersz < nrlinig1; wiersz++)
                    {
                        if (dgridzlyformat == false)
                        {
                            double a;
                            a = Convert.ToDouble(dGrid2[kolumna, wiersz].Value);

                            if (a <= 0)
                            {
                                ujemnelubzero = true;
                            }
                        }
                    }
                }


                for (int kolumna = 0; kolumna < 8; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig2; wiersz++)
                    {
                        if (dGrid3[kolumna, wiersz].Value == null)
                        {
                            brakkolumn17g2 = true;
                        }
                    }
                }
                for (int kolumna = 0; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig2; wiersz++)
                    {
                        if (dGrid3[kolumna, wiersz].Value != null)
                        {
                            string? value = dGrid3.Rows[wiersz].Cells[kolumna].Value.ToString();
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
                    for (int wiersz = 0; wiersz < nrlinig2; wiersz++)
                    {
                        if (dgridzlyformat == false)
                        {
                            double a;
                            a = Convert.ToDouble(dGrid3[kolumna, wiersz].Value);

                            if (a <= 0)
                            {
                                ujemnelubzero = true;
                            }
                        }
                    }
                }


                for (int kolumna = 0; kolumna < 8; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig3; wiersz++)
                    {
                        if (dGrid4[kolumna, wiersz].Value == null)
                        {
                            brakkolumn17g3 = true;
                        }
                    }
                }
                for (int kolumna = 0; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig3; wiersz++)
                    {
                        if (dGrid4[kolumna, wiersz].Value != null)
                        {
                            string? value = dGrid4.Rows[wiersz].Cells[kolumna].Value.ToString();
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
                    for (int wiersz = 0; wiersz < nrlinig3; wiersz++)
                    {
                        if (dgridzlyformat == false)
                        {
                            double a;
                            a = Convert.ToDouble(dGrid4[kolumna, wiersz].Value);

                            if (a <= 0)
                            {
                                ujemnelubzero = true;
                            }
                        }
                    }
                }
                if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == ""
                    || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == ""
                    || textBox7.Text == "" || textBox8.Text == "")
                {
                    braktxtboxa = true;
                }
                if (textBox1.Text == "0" || textBox2.Text == "0" || textBox3.Text == "0"
                    || textBox4.Text == "0" || textBox5.Text == "0" || textBox6.Text == "0"
                    || textBox7.Text == "0" || textBox8.Text == "0")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == "0," || textBox2.Text == "0," || textBox3.Text == "0,"
                    || textBox4.Text == "0," || textBox5.Text == "0," || textBox6.Text == "0,"
                    || textBox7.Text == "0," || textBox8.Text == "0,")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == ",0" || textBox2.Text == ",0" || textBox3.Text == ",0"
                    || textBox4.Text == ",0" || textBox5.Text == ",0" || textBox6.Text == ",0"
                    || textBox7.Text == ",0" || textBox8.Text == ",0")
                {
                    ujemnelubzero = true;
                }
                if (textBox1.Text == "," || textBox2.Text == "," || textBox3.Text == "," || textBox4.Text == ","
                           || textBox5.Text == "," || textBox6.Text == "," || textBox7.Text == "," || textBox8.Text == ",")
                {
                    txtboxzlyformat = true;
                }
            }
            if (brakkolumnzas == true || brakkolumn17g1 == true || brakkolumn17g2 == true
                || brakkolumn17g3 == true || braktxtboxa == true)
            {
                string message5 = "Brak wymaganych wartości, uzpupełnij puste komórki";
                MessageBox.Show(message5);
            }
            else if (txtboxzlyformat==true) 
            {
                string message6 = "Zły format, proszę wprowadzić liczbę";
                MessageBox.Show(message6);
            }
            else if (dgridzlyformat == true)
            {
                string message7 = "Zły format, proszę wprowadzić liczbę";
                MessageBox.Show(message7);
            }
            else if(ujemnelubzero==true)
            {
                string message8 = "Zła wartość liczby, proszę podać liczbę większą od zera";
                MessageBox.Show(message8);
            }
            else
            {
                dataGridView1.RowCount = 9;
                dataGridView2.RowCount = 10;
                dataGridView3.RowCount = 10;
                dataGridView12.RowCount = 11;


                double przewodnoscwlasc = Convert.ToDouble(textBox2.Text);
                double napznam = Convert.ToDouble(textBox3.Text) * 1000;
                double nappoczatkow = Convert.ToDouble(textBox1.Text) * 1000;
                double[] napwpkt = new double[2];
                double[] napwpktg1 = new double[nrlinig1 + 1];
                double[] napwpktg2 = new double[nrlinig2 + 1];
                double[] napwpktg3 = new double[nrlinig3 + 1];

                Complex impedancjalinizas;
                Complex impedancjatrafo;
                Complex pradrozplzas;  //prąd zasilający
                Complex spadeknapzas; // spadek napiecia na lini zasilajacej

                //gałąź 1
                double[] dlodcg1 = new double[nrlinig1];     // długość odcinka gałąź 1
                double[] przkablag1 = new double[nrlinig1];  // przekrój odcinka gałąź 1
                double[] reaktjedng1 = new double[nrlinig1]; // reaktancja odcinka gałąź 1
                double[] mocg1 = new double[nrlinig1];       // moc odbioru gałąź 1
                double[] wspmocg1 = new double[nrlinig1];    // wsp moc odbioru gałąź 1
                string[] chtrobcg1 = new string[nrlinig1];   //charakter odb gałąź 1
                double[] sinusg1 = new double[nrlinig1];     // sinus gałąź 1
                double[] cosinusg1 = new double[nrlinig1];   // cosinus gałąź 1
                Complex[] pradodbg1 = new Complex[nrlinig1]; // prady odbiorów gałąź 1
                Complex[] impedancjalinig1 = new Complex[nrlinig1];
                Complex[] pradrozplg1 = new Complex[nrlinig1 + 1];
                Complex[] spadeknapg1 = new Complex[nrlinig1];
                //gałąź 2
                double[] dlodcg2 = new double[nrlinig2];     // długość odcinka gałąź 2
                double[] przkablag2 = new double[nrlinig2];  // przekrój odcinka gałąź 2
                double[] reaktjedng2 = new double[nrlinig2]; // reaktancja odcinka gałąź 2
                double[] mocg2 = new double[nrlinig2];       // moc odbioru gałąź 2
                double[] wspmocg2 = new double[nrlinig2];    // wsp moc odbioru gałąź 2
                string[] chtrobcg2 = new string[nrlinig2];   //charakter odb gałąź 2
                double[] sinusg2 = new double[nrlinig2];     // sinus gałąź 2
                double[] cosinusg2 = new double[nrlinig2];   // cosinus gałąź 2
                Complex[] pradodbg2 = new Complex[nrlinig2]; // prady odbiorów gałąź 2
                Complex[] impedancjalinig2 = new Complex[nrlinig2];
                Complex[] pradrozplg2 = new Complex[nrlinig2 + 1];
                Complex[] spadeknapg2 = new Complex[nrlinig2];
                //gałąź 3
                double[] dlodcg3 = new double[nrlinig3];     // długość odcinka gałąź 3
                double[] przkablag3 = new double[nrlinig3];  // przekrój odcinka gałąź 3
                double[] reaktjedng3 = new double[nrlinig3]; // reaktancja odcinka gałąź 3
                double[] mocg3 = new double[nrlinig3];       // moc odbioru gałąź 3
                double[] wspmocg3 = new double[nrlinig3];    // wsp moc odbioru gałąź 3
                string[] chtrobcg3 = new string[nrlinig3];   //charakter odb gałąź 3
                double[] sinusg3 = new double[nrlinig3];     // sinus gałąź 3
                double[] cosinusg3 = new double[nrlinig3];   // cosinus gałąź 3
                Complex[] pradodbg3 = new Complex[nrlinig3]; // prady odbiorów gałąź 3
                Complex[] impedancjalinig3 = new Complex[nrlinig3];
                Complex[] pradrozplg3 = new Complex[nrlinig3 + 1];
                Complex[] spadeknapg3 = new Complex[nrlinig3];

                

                dataGridView2.Rows[0].HeaderCell.Value = "I01";
                dataGridView3.Rows[0].HeaderCell.Value = "ΔU01";

                for (int h = 0; h < nrlinig1; h++)
                {
                    dataGridView1.Rows[h].HeaderCell.Value = "I" + (h + 2);
                    dataGridView2.Rows[h + 1].HeaderCell.Value = "I" + (h + 1) + (h + 2);
                    dataGridView3.Rows[h + 1].HeaderCell.Value = "ΔU" + (h + 1) + (h + 2);
                }

                dataGridView1.Rows[3].HeaderCell.Value = "I" + (5);
                dataGridView2.Rows[4].HeaderCell.Value = "I" + (1) + (5);
                dataGridView3.Rows[4].HeaderCell.Value = "ΔU" + (1) + (5);
                for (int h = 1; h < nrlinig2; h++)
                {
                    dataGridView1.Rows[h + 3].HeaderCell.Value = "I" + (h + 5);
                    dataGridView2.Rows[h + 4].HeaderCell.Value = "I" + (h + 4) + (h + 5);
                    dataGridView3.Rows[h + 4].HeaderCell.Value = "ΔU" + (h + 4) + (h + 5);
                }

                dataGridView1.Rows[6].HeaderCell.Value = "I" + (8);
                dataGridView2.Rows[7].HeaderCell.Value = "I" + (1) + (8);
                dataGridView3.Rows[7].HeaderCell.Value = "ΔU" + (1) + (8);
                for (int h = 1; h < nrlinig3; h++)
                {
                    dataGridView1.Rows[h + 6].HeaderCell.Value = "I" + (h + 8);
                    dataGridView2.Rows[h + 7].HeaderCell.Value = "I" + (h + 7) + (h + 8);
                    dataGridView3.Rows[h + 7].HeaderCell.Value = "ΔU" + (h + 7) + (h + 8);
                }

                for (int h = 0; h < 11; h++)
                {
                    dataGridView12.Rows[h].HeaderCell.Value = "U" + h;
                }


                dataGridView1.ReadOnly = true;
                dataGridView2.ReadOnly = true;
                dataGridView3.ReadOnly = true;
                dataGridView12.ReadOnly = true;
                dataGridView4.ReadOnly = true;
                dataGridView4.AllowUserToResizeRows = false;

                //gałąź 1 - przeniesienie parametrów układu z grida do tablic
                for (int i = 0; i < nrlinig1; i++)
                {
                    dlodcg1[i] = Convert.ToDouble(dGrid2[1, i].Value);
                    przkablag1[i] = Convert.ToDouble(dGrid2[2, i].Value);
                    reaktjedng1[i] = Convert.ToDouble(dGrid2[3, i].Value);
                    mocg1[i] = Convert.ToDouble(dGrid2[5, i].Value) * 1000;

                    chtrobcg1[i] = (string)dGrid2[7, i].Value;
                }
                //gałąź 2 - przeniesienie parametrów układu z grida do tablic
                for (int i = 0; i < nrlinig2; i++)
                {
                    dlodcg2[i] = Convert.ToDouble(dGrid3[1, i].Value);
                    przkablag2[i] = Convert.ToDouble(dGrid3[2, i].Value);
                    reaktjedng2[i] = Convert.ToDouble(dGrid3[3, i].Value);
                    mocg2[i] = Convert.ToDouble(dGrid3[5, i].Value) * 1000;

                    chtrobcg2[i] = (string)dGrid3[7, i].Value;
                }
                //gałąź 3 - przeniesienie parametrów układu z grida do tablic
                for (int i = 0; i < nrlinig3; i++)
                {
                    dlodcg3[i] = Convert.ToDouble(dGrid4[1, i].Value);
                    przkablag3[i] = Convert.ToDouble(dGrid4[2, i].Value);
                    reaktjedng3[i] = Convert.ToDouble(dGrid4[3, i].Value);
                    mocg3[i] = Convert.ToDouble(dGrid4[5, i].Value) * 1000;

                    chtrobcg3[i] = (string)dGrid4[7, i].Value;
                }


                // gałąź 1 - prady na odb
                for (int h = 0; h < nrlinig1; h++)
                {
                    wspmocg1[h] = Convert.ToDouble(dGrid2[6, h].Value);
                    cosinusg1[h] = wspmocg1[h];
                    sinusg1[h] = Math.Round(Math.Sqrt(1 - Math.Pow(cosinusg1[h], 2)), 2);
                    pradodbg1[h].real = Math.Round((mocg1[h] / (Math.Sqrt(3) * napznam * cosinusg1[h])) * cosinusg1[h], 2);

                    if (chtrobcg1[h] == "Indukcyjny")
                    {
                        pradodbg1[h].imaginary = Math.Round((mocg1[h] / (Math.Sqrt(3) * napznam * cosinusg1[h])) * (sinusg1[h] * (-1)), 2);
                    }
                    else
                    {
                        pradodbg1[h].imaginary = Math.Round((mocg1[h] / (Math.Sqrt(3) * napznam * cosinusg1[h])) * sinusg1[h], 2);
                    }



                    dataGridView1[0, h].Value = pradodbg1[h].ToString();



                }

                // gałąź 2 - prady na odb
                for (int h = 0; h < nrlinig2; h++)
                {
                    wspmocg2[h] = Convert.ToDouble(dGrid3[6, h].Value);
                    cosinusg2[h] = wspmocg2[h];
                    sinusg2[h] = Math.Round(Math.Sqrt(1 - Math.Pow(cosinusg2[h], 2)), 2);
                    pradodbg2[h].real = Math.Round((mocg2[h] / (Math.Sqrt(3) * napznam * cosinusg2[h])) * cosinusg2[h], 2);

                    if (chtrobcg2[h] == "Indukcyjny")
                    {
                        pradodbg2[h].imaginary = Math.Round((mocg2[h] / (Math.Sqrt(3) * napznam * cosinusg2[h])) * (sinusg2[h] * (-1)), 2);
                    }
                    else
                    {
                        pradodbg2[h].imaginary = Math.Round((mocg2[h] / (Math.Sqrt(3) * napznam * cosinusg2[h])) * sinusg2[h], 2);
                    }



                    dataGridView1[0, h + 3].Value = pradodbg2[h].ToString();



                }

                // gałąź 3 - prady na odb
                for (int h = 0; h < nrlinig3; h++)
                {
                    wspmocg3[h] = Convert.ToDouble(dGrid4[6, h].Value);
                    cosinusg3[h] = wspmocg3[h];
                    sinusg3[h] = Math.Round(Math.Sqrt(1 - Math.Pow(cosinusg3[h], 2)), 2);
                    pradodbg3[h].real = Math.Round((mocg3[h] / (Math.Sqrt(3) * napznam * cosinusg3[h])) * cosinusg3[h], 2);

                    if (chtrobcg3[h] == "Indukcyjny")
                    {
                        pradodbg3[h].imaginary = Math.Round((mocg3[h] / (Math.Sqrt(3) * napznam * cosinusg3[h])) * (sinusg3[h] * (-1)), 2);
                    }
                    else
                    {
                        pradodbg3[h].imaginary = Math.Round((mocg3[h] / (Math.Sqrt(3) * napznam * cosinusg3[h])) * sinusg3[h], 2);
                    }



                    dataGridView1[0, h + 6].Value = pradodbg3[h].ToString();

                }

                for (int d = dataGridView1.RowCount - 2; d > -1; d--)
                {

                    if (dataGridView1.Rows[d].Cells[0].Value == null)
                    {
                        dataGridView1.Rows.RemoveAt(d);
                    }
                }
                /*if (dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0].Value == null)
                {
                    dataGridView1.Rows[dataGridView1.RowCount - 1].Visible = false;
                }*/


                double dlodc = Convert.ToDouble(dGrid1[1, 0].Value);
                double przekabla = Convert.ToDouble(dGrid1[2, 0].Value);
                double reaktjedn = Convert.ToDouble(dGrid1[3, 0].Value);
                
                
                impedancjalinizas.real = Math.Round(((1000 / (przewodnoscwlasc * przekabla)) * dlodc), 2);
                impedancjalinizas.imaginary = Math.Round(reaktjedn * dlodc, 2);
               

                if (checkBox1.Checked)
                {
                    double deltapcutrafo = Convert.ToDouble(textBox4.Text);
                    double napzwarciatrafo = Convert.ToDouble(textBox5.Text);
                    double moctrafo = Convert.ToDouble(textBox6.Text);
                    double napuzwjgornego = Convert.ToDouble(textBox7.Text);
                    double napuzwdolnego = Convert.ToDouble(textBox8.Text);
                    impedancjatrafo.real = Math.Round((deltapcutrafo * Math.Pow(napuzwdolnego, 2)) / (100 * moctrafo), 6);
                    impedancjatrafo.imaginary = Math.Round((napzwarciatrafo * Math.Pow(napuzwdolnego, 2)) / (100 * moctrafo), 6);

                    impedancjalinizas.real = impedancjatrafo.real;
                    impedancjalinizas.imaginary = impedancjatrafo.imaginary;
                }
                
                //double rezyst;
                //double reakt;
                for (int l = 0; l < nrlinig1; l++)
                {
                    impedancjalinig1[l].real = Math.Round(((1000 / (przewodnoscwlasc * przkablag1[l])) * dlodcg1[l]), 2);
                    impedancjalinig1[l].imaginary = Math.Round(reaktjedng1[l] * dlodcg1[l], 2);

                    
                }

                for (int l = 0; l < nrlinig2; l++)
                {
                    impedancjalinig2[l].real = Math.Round(((1000 / (przewodnoscwlasc * przkablag2[l])) * dlodcg2[l]), 2);
                    impedancjalinig2[l].imaginary = Math.Round(reaktjedng2[l] * dlodcg2[l], 2);

                    
                }
                for (int l = 0; l < nrlinig3; l++)
                {
                    impedancjalinig3[l].real = Math.Round(((1000 / (przewodnoscwlasc * przkablag3[l])) * dlodcg3[l]), 2);
                    impedancjalinig3[l].imaginary = Math.Round(reaktjedng3[l] * dlodcg3[l], 2);

                    
                }


                for (int h = nrlinig1 - 1; h > -1; h--)
                {
                    if (h == nrlinig1 - 1)
                    {
                        pradrozplg1[h + 1].real = 0;
                        pradrozplg1[h + 1].imaginary = 0;
                    }
                    pradrozplg1[h].real = Math.Round(pradodbg1[h].real + pradrozplg1[h + 1].real, 2);
                    pradrozplg1[h].imaginary = Math.Round(pradodbg1[h].imaginary + pradrozplg1[h + 1].imaginary, 2);

                }
                for (int h = 0; h < nrlinig1; h++)
                {
                    dataGridView2[0, h + 1].Value = pradrozplg1[h].ToString();
                }


                for (int h = nrlinig2 - 1; h > -1; h--)
                {
                    if (h == nrlinig2 - 1)
                    {
                        pradrozplg2[h + 1].real = 0;
                        pradrozplg2[h + 1].imaginary = 0;
                    }
                    pradrozplg2[h].real = Math.Round(pradodbg2[h].real + pradrozplg2[h + 1].real, 2);
                    pradrozplg2[h].imaginary = Math.Round(pradodbg2[h].imaginary + pradrozplg2[h + 1].imaginary, 2);

                }
                for (int h = 0; h < nrlinig2; h++)
                {
                    dataGridView2[0, h + 4].Value = pradrozplg2[h].ToString();
                }


                for (int h = nrlinig3 - 1; h > -1; h--)
                {
                    if (h == nrlinig3 - 1)
                    {
                        pradrozplg3[h + 1].real = 0;
                        pradrozplg3[h + 1].imaginary = 0;
                    }
                    pradrozplg3[h].real = Math.Round(pradodbg3[h].real + pradrozplg3[h + 1].real, 2);
                    pradrozplg3[h].imaginary = Math.Round(pradodbg3[h].imaginary + pradrozplg3[h + 1].imaginary, 2);

                }
                for (int h = 0; h < nrlinig3; h++)
                {
                    dataGridView2[0, h + 7].Value = pradrozplg3[h].ToString();
                }

                pradrozplzas.real = Math.Round(pradrozplg1[0].real + pradrozplg2[0].real + pradrozplg3[0].real, 2);
                pradrozplzas.imaginary = Math.Round(pradrozplg1[0].imaginary + pradrozplg2[0].imaginary + pradrozplg3[0].imaginary, 2);
                dataGridView2[0, 0].Value = pradrozplzas.ToString();
                int nrlini = (int)numericUpDown1.Value;
                dataGridView4.RowCount = 2;
                if (checkBox1.Checked) 
                {
                    dataGridView4.Visible = true;
                    double napuzwjgornego = Convert.ToDouble(textBox7.Text);
                    double napuzwdolnego = Convert.ToDouble(textBox8.Text);
                    Complex pradwysoki = new Complex(0, 0);
                    pradwysoki = pradrozplzas * ( napuzwdolnego/ napuzwjgornego);
                    pradwysoki.real = Math.Round(pradwysoki.real,2);
                    pradwysoki.imaginary = Math.Round(pradwysoki.imaginary, 2);
                    dataGridView4.Rows[0].HeaderCell.Value = "Iz0" + "(" + napuzwdolnego + ")";
                    dataGridView4.Rows[1].HeaderCell.Value = "Iz0" + "(" + napuzwjgornego + ")";
                    dataGridView4[0, 0].Value = pradrozplzas;
                    dataGridView4[0, 1].Value = pradwysoki;
                }


                for (int d = dataGridView2.RowCount - 2; d > -1; d--)
                {

                    if (dataGridView2.Rows[d].Cells[0].Value == null)
                    {
                        dataGridView2.Rows.RemoveAt(d);
                    }
                }
                if (dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[0].Value == null)
                {
                    dataGridView2.Rows[dataGridView2.RowCount - 1].Visible = false;
                }



                for (int c = 0; c < nrlinig1; c++)
                {

                    spadeknapg1[c] = pradrozplg1[c] * impedancjalinig1[c];
                    spadeknapg1[c].real = Math.Round(spadeknapg1[c].real * Math.Sqrt(3), 2);
                    spadeknapg1[c].imaginary = Math.Round(spadeknapg1[c].imaginary * Math.Sqrt(3), 2);

                    dataGridView3[0, c + 1].Value = spadeknapg1[c].real.ToString();


                }

                for (int c = 0; c < nrlinig2; c++)
                {

                    spadeknapg2[c] = pradrozplg2[c] * impedancjalinig2[c];
                    spadeknapg2[c].real = Math.Round(spadeknapg2[c].real * Math.Sqrt(3), 2);
                    spadeknapg2[c].imaginary = Math.Round(spadeknapg2[c].imaginary * Math.Sqrt(3), 2);

                    dataGridView3[0, c + 4].Value = spadeknapg2[c].real.ToString();


                }


                for (int c = 0; c < nrlinig3; c++)
                {

                    spadeknapg3[c] = pradrozplg3[c] * impedancjalinig3[c];
                    spadeknapg3[c].real = Math.Round(spadeknapg3[c].real * Math.Sqrt(3), 2);
                    spadeknapg3[c].imaginary = Math.Round(spadeknapg3[c].imaginary * Math.Sqrt(3), 2);

                    dataGridView3[0, c + 7].Value = spadeknapg3[c].real.ToString();


                }


                spadeknapzas = pradrozplzas * impedancjalinizas;
                spadeknapzas.real = Math.Round(spadeknapzas.real * Math.Sqrt(3), 2);
                spadeknapzas.imaginary = Math.Round(spadeknapzas.imaginary * Math.Sqrt(3), 2);
                dataGridView3[0, 0].Value = spadeknapzas.real.ToString();

                for (int d = dataGridView3.RowCount - 2; d > -1; d--)
                {

                    if (dataGridView3.Rows[d].Cells[0].Value == null)
                    {
                        dataGridView3.Rows.RemoveAt(d);
                    }
                }
                if (dataGridView3.Rows[dataGridView3.RowCount - 1].Cells[0].Value == null)
                {
                    dataGridView3.Rows[dataGridView3.RowCount - 1].Visible = false;
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
               
                napwpkt[1] = Math.Round(napwpkt[0] - spadeknapzas.real, 2);

                napwpktg1[0] = Math.Round(napwpkt[1] - spadeknapg1[0].real, 2);
                napwpktg2[0] = Math.Round(napwpkt[1] - spadeknapg2[0].real, 2);
                napwpktg3[0] = Math.Round(napwpkt[1] - spadeknapg3[0].real, 2);

                for (int d = 1; d < nrlinig1; d++)
                {
                    napwpktg1[d] = napwpktg1[d - 1] - spadeknapg1[d].real;
                   
                }

                for (int d = 1; d < nrlinig2; d++)
                {
                    napwpktg2[d] = napwpktg2[d - 1] - spadeknapg2[d].real;
                   
                }

                for (int d = 1; d < nrlinig3; d++)
                {
                    napwpktg3[d] = napwpktg3[d - 1] - spadeknapg3[d].real;
                    
                }


                dataGridView12[0, 0].Value = napwpkt[0].ToString();
                dataGridView12[0, 1].Value = napwpkt[1].ToString();

                for (int d = 0; d < nrlinig1; d++)
                {
                    napwpktg1[d] = Math.Round(napwpktg1[d], 2);

                    dataGridView12[0, d + 2].Value = napwpktg1[d].ToString();


                }

                for (int d = 0; d < nrlinig2; d++)
                {

                    napwpktg2[d] = Math.Round(napwpktg2[d], 2);
                    dataGridView12[0, d + 5].Value = napwpktg2[d].ToString();

                }

                for (int d = 0; d < nrlinig3; d++)
                {

                    napwpktg3[d] = Math.Round(napwpktg3[d], 2);
                    dataGridView12[0, d + 8].Value = napwpktg3[d].ToString();

                }

                for (int d = dataGridView12.RowCount - 2; d > -1; d--)
                {

                    if (dataGridView12.Rows[d].Cells[0].Value == null)
                    {
                        dataGridView12.Rows.RemoveAt(d);
                    }
                }
                if (dataGridView12.Rows[dataGridView12.RowCount - 1].Cells[0].Value == null)
                {
                    dataGridView12.Rows[dataGridView12.RowCount - 1].Visible = false;
                }


            }
        }
        private void button_przyklad1_Click(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory();
            string fullpath = path + "\\" + "przyklady\\przyklad1r.txt";
            string text = File.ReadAllText(fullpath);


            string[] words = text.Split(new char[] {' ','\n'});
            textBox1.Text = Convert.ToString(words[0]);
            textBox3.Text = Convert.ToString(words[1]);
            textBox2.Text = Convert.ToString(words[2]);

            int nrlinig1 = Convert.ToInt32(words[3]);
            int nrlinig2 = Convert.ToInt32(words[4]);
            int nrlinig3 = Convert.ToInt32(words[5]);

            dGrid1[1, 0].Value = words[6];
            dGrid1[2, 0].Value = words[7];
            dGrid1[3, 0].Value = words[8];


            string wspg11 = "Indukcyjny";
            string wspg12 = "Indukcyjny";
            string wspg21 = "Pojemnościowy";
            string wspg31 = "Indukcyjny";


            numericUpDown1.Value = nrlinig1;
            numericUpDown2.Value = nrlinig2;
            numericUpDown3.Value = nrlinig3;

            for (int kolumna = 0; kolumna < 7; kolumna++)
            {
                for (int wiersz = 0; wiersz < nrlinig1; wiersz++)
                {
                    dGrid2[kolumna, wiersz].Value = words[9 + wiersz * 7 + kolumna];
                    

                }
            }
            for (int kolumna = 0; kolumna < 7; kolumna++)
            {
                for (int wiersz = 0; wiersz < nrlinig2; wiersz++)
                {
                    dGrid3[kolumna, wiersz].Value = words[23 + wiersz * 7 + kolumna];


                }
            }
            for (int kolumna = 0; kolumna < 7; kolumna++)
            {
                for (int wiersz = 0; wiersz < nrlinig3; wiersz++)
                {
                    dGrid4[kolumna, wiersz].Value = words[30 + wiersz * 7 + kolumna];


                }
            }
            dGrid2[7, 0].Value = wspg11;
            dGrid2[7, 1].Value = wspg12;
            dGrid3[7, 0].Value = wspg21;
            dGrid4[7, 0].Value = wspg31;
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DataGridViewComboBoxColumn columng1 = new DataGridViewComboBoxColumn();
            columng1.HeaderText = "Ch-ter obciążenia";
            columng1.Name = " "; 
            columng1.Items.AddRange("Indukcyjny", "Pojemnościowy");
            DataGridViewComboBoxColumn columng2 = new DataGridViewComboBoxColumn();
            columng2.HeaderText = "Ch-ter obciążenia";
            columng2.Name = " ";
            columng2.Items.AddRange("Indukcyjny", "Pojemnościowy");
            DataGridViewComboBoxColumn columng3 = new DataGridViewComboBoxColumn();
            columng3.HeaderText = "Ch-ter obciążenia";
            columng3.Name = " ";
            columng3.Items.AddRange("Indukcyjny", "Pojemnościowy");
            
            numericUpDown1.Value = 1;
            numericUpDown2.Value = 1;
            numericUpDown3.Value = 1;
            for (int i = 1; i < 4; i++)
            {
                dGrid1[i, 0].Value = null;
                dGrid2[i, 0].Value = null;
                dGrid3[i, 0].Value = null;
                dGrid4[i, 0].Value = null;

            }
            for (int i = 5; i < 7; i++)
            {
                dGrid2[i, 0].Value = null;
                dGrid3[i, 0].Value = null;
                dGrid4[i, 0].Value = null;
            }
           
            
            dGrid2.Columns.RemoveAt(7);
            dGrid3.Columns.RemoveAt(7);
            dGrid4.Columns.RemoveAt(7);
            dGrid2.Columns.Add(columng1);
            dGrid3.Columns.Add(columng2);
            dGrid4.Columns.Add(columng3);
            checkBox1.Checked = false;
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            dataGridView1.RowCount = 1;
            dataGridView1[0, 0].Value = null;
            dataGridView1.Rows[0].HeaderCell.Value = "";
            dataGridView2.RowCount = 1;
            dataGridView2[0, 0].Value = null;
            dataGridView2.Rows[0].HeaderCell.Value = "";
            dataGridView3.RowCount = 1;
            dataGridView3[0, 0].Value = null;
            dataGridView3.Rows[0].HeaderCell.Value = "";
            dataGridView12.RowCount = 1;
            dataGridView12[0, 0].Value = null;
            dataGridView12.Rows[0].HeaderCell.Value = "";
            dataGridView4.RowCount = 1;
            dataGridView4[0, 0].Value = null;
            dataGridView4.Rows[0].HeaderCell.Value = "";

            pictureBox1.Image = null;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int iloscodinkow1 = (int)numericUpDown1.Value;
            int iloscodinkow2 = (int)numericUpDown2.Value;
            int iloscodinkow3 = (int)numericUpDown3.Value;

            string odbnr;
            string path = Directory.GetCurrentDirectory();

            string fullpath = path + "\\" + "obrazki\\";
            if (!checkBox1.Checked)
            {
                odbnr = "r"+iloscodinkow1+iloscodinkow2+iloscodinkow3;
                Image image = Image.FromFile(fullpath + odbnr + ".jpg");
                this.pictureBox1.Image = image;
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            }
            else if (checkBox1.Checked)
            {
                odbnr = "r"+iloscodinkow1+iloscodinkow2+iloscodinkow3+"t";
                Image image = Image.FromFile(@fullpath + odbnr + ".jpg");
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
            dataGridView12.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView12.MultiSelect = true;
            dataGridView12.SelectAll();
            DataObject dataObj4 = dataGridView12.GetClipboardContent();
            if (dataObj4 != null)
                Clipboard.SetDataObject(dataObj4);

        }
        private void copyAlltoClipboard5()
        {
            dataGridView4.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView4.MultiSelect = true;
            dataGridView4.SelectAll();
            DataObject dataObj5 = dataGridView4.GetClipboardContent();
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

        private void dGrid1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
           /* if (!checkBox1.Checked)
            {
                for (int kolumna = 0; kolumna < 4; kolumna++)
                {
                    for (int wiersz = 0; wiersz < 1; wiersz++)
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

        private void dGrid2_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            /*int nrlinig1 = (int)numericUpDown1.Value;
            if (!checkBox1.Checked || checkBox1.Checked)
            {
                for (int kolumna = 1; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig1; wiersz++)
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

        private void dGrid3_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
           /* int nrlinig2 = (int)numericUpDown2.Value;
            if (!checkBox1.Checked || checkBox1.Checked)
            {
                for (int kolumna = 1; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig2; wiersz++)
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

        private void dGrid4_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
           /* int nrlinig3 = (int)numericUpDown3.Value;
            if (!checkBox1.Checked || checkBox1.Checked)
            {
                for (int kolumna = 1; kolumna < 7; kolumna++)
                {
                    for (int wiersz = 0; wiersz < nrlinig3; wiersz++)
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

        private void button_przyklad2_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = true;
            string path = Directory.GetCurrentDirectory();
            string fullpath = path + "\\" + "przyklady\\przyklad2r.txt";
            string text1 = File.ReadAllText(fullpath);

            string[] words1 = text1.Split(new char[] { ' ', '\n' });
            textBox1.Text = Convert.ToString(words1[0]);
            textBox3.Text = Convert.ToString(words1[1]);
            textBox2.Text = Convert.ToString(words1[2]);
            textBox4.Text = Convert.ToString(words1[6]);
            textBox5.Text = Convert.ToString(words1[7]);
            textBox6.Text = Convert.ToString(words1[8]);
            textBox7.Text = Convert.ToString(words1[9]);
            textBox8.Text = Convert.ToString(words1[10]);

            int nrlinig1 = Convert.ToInt32(words1[3]);
            int nrlinig2 = Convert.ToInt32(words1[4]);
            int nrlinig3 = Convert.ToInt32(words1[5]);

            string wspg11 = "Indukcyjny";
            string wspg12 = "Indukcyjny";
            string wspg21 = "Pojemnościowy";
            string wspg31 = "Indukcyjny";


            numericUpDown1.Value = nrlinig1;
            numericUpDown2.Value = nrlinig2;
            numericUpDown3.Value = nrlinig3;

            for (int kolumna = 0; kolumna < 7; kolumna++)
            {
                for (int wiersz = 0; wiersz < nrlinig1; wiersz++)
                {
                    dGrid2[kolumna, wiersz].Value = words1[11 + wiersz * 7 + kolumna];


                }
            }
            for (int kolumna = 0; kolumna < 7; kolumna++)
            {
                for (int wiersz = 0; wiersz < nrlinig2; wiersz++)
                {
                    dGrid3[kolumna, wiersz].Value = words1[25 + wiersz * 7 + kolumna];


                }
            }
            for (int kolumna = 0; kolumna < 7; kolumna++)
            {
                for (int wiersz = 0; wiersz < nrlinig3; wiersz++)
                {
                    dGrid4[kolumna, wiersz].Value = words1[32 + wiersz * 7 + kolumna];


                }
            }
            dGrid2[7, 0].Value = wspg11;
            dGrid2[7, 1].Value = wspg12;
            dGrid3[7, 0].Value = wspg21;
            dGrid4[7, 0].Value = wspg31;
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }
    }
}
