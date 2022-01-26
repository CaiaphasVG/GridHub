using System;
using System.Windows.Forms;
using System.Threading;


namespace NEA_Prototype
{
    public partial class Form2 : Form
    {

        Thread th; //Thread instance to open the main menu
        CreateUser createUser = new CreateUser(); //Instanciates a new instance of the CreateUser class

        public Form2()
        {
            InitializeComponent();
            maskedTextBox4.PasswordChar = '*';
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox4_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox4_Click(object sender, EventArgs e)
        {
            //Clears the textbox and sets the password character
            maskedTextBox4.Text = "";
        }

        private void maskedTextBox5_Click(object sender, EventArgs e)
        {
            //Clears the textbox
            maskedTextBox5.Text = "";
        }

        private void maskedTextBox1_Click(object sender, EventArgs e)
        {
            //Clears the textbox
            maskedTextBox1.Text = "";
        }

        private void maskedTextBox2_Click(object sender, EventArgs e)
        {
            //Clears the textbox
            maskedTextBox2.Text = "";
        }

        private void maskedTextBox3_Click(object sender, EventArgs e)
        {
            //Clears the textbox
            maskedTextBox3.Text = "";
        }

        //Checks the length and validity of all entered credentials
        private void button1_Click(object sender, EventArgs e)
        {
            //Sets the local variable correctCred to 0, with each valid entry it increments to one
            int correctCred = 0;

            //Each textbox input is checked by the corresponding 
            createUser.CheckFirstName(maskedTextBox1.Text);
            if (createUser.isValid == false)
                MessageBox.Show("The first name was invalid, make sure the textbox contains no numbers", "Invalid detail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                correctCred += 1;
            createUser.CheckLastName(maskedTextBox5.Text);
            if (createUser.isValid == false)
                MessageBox.Show("The last name was invalid, make sure the textbox contains no numbers", "Invalid detail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                correctCred += 1;
            createUser.CheckUsername(maskedTextBox2.Text);
            if (createUser.isValid == false)
                MessageBox.Show("The username was invalid, make sure the username is 3 to 20 characters in length", "Invalid detail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                correctCred += 1;
            createUser.CheckRole(maskedTextBox3.Text);
            if (createUser.isValid == false)
                MessageBox.Show("The role was invalid, make sure the textbox contains no numbers", "Invalid detail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                correctCred += 1;
            createUser.CheckPassword(maskedTextBox4.Text);
            if (createUser.isValid == false)
                MessageBox.Show("The password was invalid, make sure the password is at least 6 characters in length, cotains a number and a upper case character", "Invalid detail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                correctCred += 1;

            //Check whether we got 5 correct inputs 
            if (correctCred == 5)
            {
                MessageBox.Show("User created", "Sign up was successful!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //Sets encryption variable to true, which means encrypting
                createUser.isEncrypting = true;
                //Encrypts the user file
                createUser.ReadWriteUser("", "");
                this.Close();
                th = new Thread(openForm3);
                th.SetApartmentState(ApartmentState.STA);
                th.Start();
            }
        }

        private void openForm3(object obj)
        {
            //Opens the main menu
            Form3 form3 = new Form3();
            //Assigns the form3 user with the current user instance
            form3.user = createUser;
            //Calls method to set display name
            form3.SetName();
            //Opens main menu form
            Application.Run(new Form3());
        }

        private void maskedTextBox5_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void maskedTextBox2_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }
    }
}
