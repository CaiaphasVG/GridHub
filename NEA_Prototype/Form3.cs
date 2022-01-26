using System;
using System.Windows.Forms;
using System.Threading;


namespace NEA_Prototype
{
    public partial class Form3 : Form
    {

        Thread th;  //Thread to open different menus
        Report report = new Report();   //Initialises new instance of the report class
        int reportType; //Integer that specifies what report type is being generated
        string period;  //Period to enter into the report
        string settlementDate;  //Settlment date to enter into the report
        public CreateUser user = new CreateUser();  //New instance of the user class
        private bool validDate = false; //The variable that will store wheteher the entered date was valid
        private bool validLengths = true;   //The variable that will store whether the length entered was valid
        public DateTime date = new DateTime();  //New instance of datetime, which is used to record the time a report was generated

        public Form3()
        {
            InitializeComponent();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Stores the report type by the index of the selected combo box
            reportType = comboBox1.SelectedIndex;              
        }

        //The search button
        private void button1_Click(object sender, EventArgs e)
        {
            //Initialises the valid date to false, to avoid it being left to set to true
            validDate = false;
            //Calls the length check function
            LengthCheck(maskedTextBox2.Text, maskedTextBox3.Text);
            //If the date and length is valid then start generation process
            if(validDate == true & validLengths == true)
            {
                this.Close();
                th = new Thread(openForm5);
                th.SetApartmentState(ApartmentState.STA);
                th.Start();
            }

        }

        private void maskedTextBox2_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {
            period = maskedTextBox2.Text;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        //Log out button


        private void button3_Click(object sender, EventArgs e)
        {

        }

        //Calls to open the saved reports form
        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
            th = new Thread(openForm4);
            th.SetApartmentState(ApartmentState.STA);
            th.Start();
        }

        //Opens the saved reports form
        private void openForm4(object obj)
        {
            Form4 form4 = new Form4();
            form4.user = user;
            Application.Run(form4);
        }

        //Opens the report page
        private void openForm5(object obj) 
        {
            Form5 form5 = new Form5();
            //Report is generated 
            report.GenerateReport(reportType, settlementDate, period);
            //Sets the datatable in the report class to form 5's data table
            form5.DisplayReport(report);
            //Sets the users instance
            form5.user = user;
            Application.Run(form5);
        }

        private void LengthCheck(string period, string date)
        {
            //Checks whether period is greater than 0, whether it is equal to null and if it contains any blank spaces
            if (period.ToString() == "00" || period == null || period.Length < 2 ||date.Contains(" "))
            {
                //Sets valid length to false
                validLengths = false;
                MessageBox.Show("The period must be greater than 0 and no empty values can be entered", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                validLengths = true;
                //Once it's been establised that the length is long enough the date check is called 
                DateCheck(maskedTextBox3.Text);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            th = new Thread(openForm1);
            th.SetApartmentState(ApartmentState.STA);
            th.Start();
        }

        private void openForm1(object obj)
        {
            Application.Run(new Form1());
        }

        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox3_TextChanged(object sender, EventArgs e) 
        {

        }

        private void DateCheck(string textBox)
        {
            //Get instance of current date
            date = DateTime.Now;

            //Checks whether the it's a leap year and if it's the 29th of febuary
            if (Int16.Parse(textBox.Substring(8, 2)) == 29 & Int16.Parse(textBox.Substring(0, 4)) % 4 != 0 & Int16.Parse(textBox[5].ToString()) == 2)
            {
                MessageBox.Show("Please enter a valid date", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            //Checks whether the inputted date is greater than the current date 
            else if(Int16.Parse(textBox.Substring(0, 4)) > date.Year || ((Int16.Parse(textBox.Substring(0, 4)) == date.Year & Int16.Parse(textBox.Substring(5, 2)) > date.Month)) || ((Int16.Parse(textBox.Substring(0, 4)) == date.Year & Int16.Parse(textBox.Substring(5, 2)) == date.Month & Int16.Parse(textBox.Substring(8, 2)) > date.Day)))
            {
                MessageBox.Show("The date cannot be greater than the current date", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            //Checks whether the month inputted is between 1 and 12
            else if (Int16.Parse(textBox.Substring(5, 2)) > 12 || Int16.Parse(textBox.Substring(5, 2)) < 1)
            {
                MessageBox.Show("Please enter a valid date", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            //Checks whether the day entered is greater than 31, if the month divided by two gives a remainder
            else if (Int16.Parse(textBox.Substring(8, 2)) > 31 & Int16.Parse(textBox.Substring(5, 2)) % 2 != 0 || Int16.Parse(textBox.Substring(8, 2)) < 1)
            {
                MessageBox.Show("Please enter a valid date", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            //Checks whether the day entered is greater than 30, if the month divided by two gives a remainder
            else if (Int16.Parse(textBox.Substring(8, 2)) > 30 & Int16.Parse(textBox.Substring(5, 2)) % 2 == 0 || Int16.Parse(textBox.Substring(8, 2)) < 1)
            {
                MessageBox.Show("Please enter a valid date", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                //Sets the date
                settlementDate = textBox;
                validDate = true;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void maskedTextBox2_TextChanged(object sender, EventArgs e) 
        {

        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        public void SetName() {
            //Reads the CreateUser instance to display the name of the user
            label1.Text = "Welcome back, " + user.e_firstName;
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
