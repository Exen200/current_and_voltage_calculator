using System.Security.Cryptography.X509Certificates;

namespace WinFormsApp1
{
    public struct Complex
    {

        public double real;
        public double imaginary;


        public Complex(double real, double imaginary)
        {
            this.real = real;
            this.imaginary = imaginary;
        }
        public static Complex operator +(Complex one, Complex two)
        {
            return new Complex(one.real + two.real, one.imaginary + two.imaginary);
        }
        public static Complex operator -(Complex one, Complex two)
        {
            return new Complex(one.real - two.real, one.imaginary - two.imaginary);
        }
        public static Complex operator *(Complex one, Complex two)
        {
            return new Complex(one.real * two.real - one.imaginary * two.imaginary, one.real * two.imaginary + one.imaginary * two.real);
        }
        public static Complex operator *(Complex one, double two)
        {
            return new Complex(one.real * two, one.imaginary * two);
        }
        public static Complex operator /(Complex one, double two)
        {
            return new Complex(one.real / two, one.imaginary / two);
        }
        public static Complex operator /(Complex one, Complex two)
        {
            return new Complex((one.real * two.real + one.imaginary * two.imaginary) / (Math.Pow(two.real, 2) + Math.Pow(two.imaginary, 2)), (one.imaginary * two.real - one.real * two.imaginary) / (Math.Pow(two.real, 2) + Math.Pow(two.imaginary, 2)));
        }

        public override string ToString()
        {
            if (imaginary >= 0)
            {
                return (String.Format("{0} + {1}i", real, imaginary));
            }
            else
            {   
                return (String.Format("{0}  {1}i", real, imaginary));
            }
        }


    }
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            userControl11.Hide();
            userControl21.Hide();
            userControl31.Hide();
        }


        private void userControl11_Load(object sender, EventArgs e)
        {

        }

        private void userControl21_Load(object sender, EventArgs e)
        {

        }

        private void userControl31_Load(object sender, EventArgs e)
        {

        }
        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            string uklad = comboBox1.Text.ToString();
            if (uklad == "Uk³ad jednostronny")
            {
                userControl21.Hide();
                userControl31.Hide();
                userControl11.Show();
                userControl11.BringToFront();
            }
            else if (uklad == "Uk³ad dwustronny")
            {
                userControl11.Hide();
                userControl31.Hide();
                userControl21.Show();
                userControl21.BringToFront();
            }
            else
            {
                userControl21.Hide();
                userControl31.Hide();
                userControl31.Show();
                userControl31.BringToFront();
            }
        
        }

        private void userControl31_Load_1(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Program oblicza rozp³ywy pr¹dów oraz spadki napiêcia na podstawie wybranego uk³adu oraz danych wejœciowych.\nLista rozwijana u góry okna aplikacji pozwala okreœliæ typ badanego uk³adu.\nAby program móg³ zacz¹æ obliczaæ nale¿y wype³niæ wszystkie wymagane dane.");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("PRACA IN¯YNIERSKA\nProgram do obliczania rozp³ywów pr¹dów oraz spadków napiêæ w okreœlonych uk³adach sieci energetycznych\nWykonali: Konrad Zaj¹czkowski, Mateusz Diaków");
        }
    }
}