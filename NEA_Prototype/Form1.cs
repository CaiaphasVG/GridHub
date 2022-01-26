using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;

namespace NEA_Prototype
{
    public partial class Form1 : Form
    {

        Thread th;  //Thread instance to open the main menu and the sign up screen
        CreateUser user = new CreateUser(); //Instanciates a new instance of the CreateUser class

        public Form1()
        {
            InitializeComponent();
            maskedTextBox2.PasswordChar = '*';

        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {
            
        }

        private void maskedTextBox1_Click(object sender, EventArgs e)
        {
            //On click sets the username textbox to empty
            maskedTextBox1.Text = ""; 
        }

        private void maskedTextBox2_Click(object sender, EventArgs e) 
        {
            //On click sets the password textbox to empty and sets the password character 
            maskedTextBox2.Text = "";
        }

        private void maskedTextBox2_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {
            //Opens the sign up menu when the sign up label is clicked
            this.Close();
            th = new Thread(openForm2);
            th.SetApartmentState(ApartmentState.STA);
            th.Start();       
        }

        private void openForm2(object obj) 
        {
            Application.Run(new Form2());
        }

        private void label3_MouseHover(object sender, EventArgs e) 
        {
            //Changes the colour of the sign up text when hovered
            label3.ForeColor = Color.Blue;
        }

        private void label3_MouseLeave(object sender, EventArgs e) 
        {
            //Resets the colour when the mouse isn't hovering over
            label3.ForeColor = Color.FromArgb(215, 247, 229);
        }

        //Decrypts the usefile and checks the credentials
        private void button1_Click(object sender, EventArgs e)
        {
            //Set user encrypting to false which means decrypting
            user.isEncrypting = false;
            //Inputs the username and password inputs into the encryption method
            user.ReadWriteUser(maskedTextBox1.Text, maskedTextBox2.Text);
            //Checks if the login was a success
            if (user.validLogin == true)
            {
                //If the login was successful open the main menue
                th = new Thread(openForm3);
                this.Close();
                th.SetApartmentState(ApartmentState.STA);
                th.Start();
            }
            else
            {
                //Check whether there was a file not found error, which means there was no file found for that corresponding username which means the username was incorrect
                if (user.flEx != null)
                    MessageBox.Show("Your login was not recognised, try again.", "Invalid Login",  MessageBoxButtons.OK, MessageBoxIcon.Error);
                //We found a username file but the password was incorrect
                else
                    MessageBox.Show("Your password was not recognised, try again.", "Invalid Login", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void openForm3(object obj)
        {
            //Opens the main menu
            Form3 form3 = new Form3();
            //Assigns the form3 user with the current user instance
            form3.user = user;
            //Calls method to set display name
            form3.SetName();
            Application.Run(form3);
        }
    }
}
