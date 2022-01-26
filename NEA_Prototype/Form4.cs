using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Threading;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Data.OleDb;
using System.IO;
using Microsoft.VisualBasic;
using System.Diagnostics;


namespace NEA_Prototype
{

    
    public partial class Form4 : Form
    {
        public CreateUser user = new CreateUser();  //New instance of the user class
        Thread th;  //Thread to navigate forms
        public DataTable dataTable = new DataTable();   //Datatable to display opened report
        public DataTable dataTableMB = new DataTable();
        public DataTable dataTableBackUp = new DataTable();
        OpenFileDialog fPath = new OpenFileDialog();    //File dialouge to obtain the path of the database
        string directoryPathMDB;   //The file path
        string dbSource;    //Stores the data source and file name of the file to access
        string dbProvider;   //The microsoft provider
        OleDbConnection dbMBCon = new System.Data.OleDb.OleDbConnection();  //New connection established
        string CommandText; //String that stores the programs sql statement to write and display reports
        private bool validSaveName = false; //Variable that stores whether the save name was valid
        private List<String> illegalCharacters = new List<string>() { "<", ">", ":", Char.ConvertFromUtf32(34).ToString(), @"\", @"/", "|", "?", "*" }; //List of illegal file characters windows will not allow files to be called
        string fileNamePDF;    //Stores the user inputted name for the file to save to
        string directoryPath;   //Stores the user inputted path 
        string sqlStatement;    //The user inputted SQL statement 
        FolderBrowserDialog folderPath = new FolderBrowserDialog();  //Allows the user to select a file location to save to



        public Form4()
        {
            InitializeComponent();
            //Sets the buttons that export to PDF and SQL statements to disabled until a report is loaded
            button2.Enabled = false;
            button4.Enabled = false;
            textBox1.Enabled = false;
            button5.Enabled = false;
        }

        //Button that goes back to the main menu
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
            th = new Thread(openForm3);
            th.SetApartmentState(ApartmentState.STA);
            th.Start();
        }

        private void openForm3(object obj)
        {
            Application.Run(new Form3());
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }

        //Button that opens a file
        private void button1_Click(object sender, EventArgs e)
        {
            dataTableMB.Clear();    //Clears the current data table of any contents
            dataGridView1.DataSource = null;    //Clears the data grid
            OleDbConnection dbCon = new System.Data.OleDb.OleDbConnection();
            DialogResult result = fPath.ShowDialog();   //Opens the file dialouge
            directoryPathMDB = fPath.FileName; //Stores the file path
            if(fPath.SafeFileName.Substring(fPath.SafeFileName.Length-3) != "mdb")  //Checks whether the last 3 digits of the file name end in mdb
            {
                //Error pops up if an invalid file type is opened
                MessageBox.Show("Please open files ending in .mdb", "Invalid file type", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                dbSource = "Data source = " + Environment.CurrentDirectory + @"\\MADB.mdb"; //The masterdatabase file location is stored in dbSource
                dbProvider = "Provider=Microsoft.Jet.OLEDB.4.0;";   //The OLEdb provider needed to access the file
                dbMBCon.ConnectionString = dbProvider + dbSource;   //Compiled full connection string 
                dbMBCon.Open(); //Opening the connection
                CommandText = @"SELECT * FROM FILERECORDS WHERE [File Name] = '" + fPath.SafeFileName.Substring(0, fPath.SafeFileName.Length - 4) + "';";   //String of command to select all columns where the file name of the opened file matches the row in the database
                OleDbDataAdapter dtMB = new OleDbDataAdapter(CommandText, dbMBCon); //New adapter instance to write results to a datatable the program searches
                dtMB.Fill(dataTableMB); //Fill data table with results
                dbMBCon.Close();    //Close connection

                //Checking whether the private status is set to false, OR whether it is set to true and the currently logged in user matches the creator of the file
                if (dataTableMB.Rows[0][6].ToString() == "false" || (dataTableMB.Rows[0][6].ToString() == "true" & dataTableMB.Rows[0][1].ToString() == user.e_username))
                {
                    string conString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + directoryPathMDB + ";";   //Assigns connection string with provider and datasourse
                    dbCon.ConnectionString = conString;
                    dbCon.Open();   //Opens connection to file
                    string comString = "SELECT * FROM " + dataTableMB.Rows[0][2].ToString();   //Command string varies on the report
                    OleDbDataAdapter dt = new OleDbDataAdapter(comString, dbCon);   //New adapter instance to write results to a datatable the program searches
                    dt.Fill(dataTable); //Data table filled with returned results
                    dataGridView1.DataSource = dataTable;   //Datagridview source assignemed with data table
                    dataTableBackUp = dataTable;    //Original, unquereied data table backed up
                    //Enables exporting and SQL buttons
                    button2.Enabled = true;
                    button4.Enabled = true;
                    textBox1.Enabled = true;
                    button5.Enabled = true;
                }
                else
                {
                    MessageBox.Show("You do not have access to this file.", "Invalid permissions", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void CheckValidPDFFileName()
        {
            //Check whether the PDF file name is between 2 and 25 characters
            if (textBox1.Text.Length >= 2 & textBox1.Text.Length <= 25)
            {
                //Cycles through each element in the illegal characters list
                foreach (string illegalCharacter in illegalCharacters)
                {   
                    //Checks whether the textbox contains and illegal characters
                    if (textBox1.Text.Contains(illegalCharacter))
                    {
                        MessageBox.Show(@"The file name cannot contain any of the following characters '< , >, :, \, /, |, *, ?' or a blank space.", "Invalid name", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //Resets the text box
                        textBox1.Text = null;
                        //Sets the valid save name to false because it had an illegal character in it
                        validSaveName = false;
                        //Breaks for loop
                        break;
                    }
                    else
                        validSaveName = true;
                }
                //If the save name is valid and the directory path is valid the PDF is exported
                if (validSaveName == true)
                    if (directoryPath != null)
                        //Contains the directory path, file name and the personal information to store in the header
                        ExportPDF(dataTable, (directoryPath + "\\" + fileNamePDF + ".pdf"), ("Report Author: " + user.e_firstName + " " + user.e_lastName + " " + "Department: " + user.e_role));
                    else
                        MessageBox.Show("File location cannot be blank.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
                MessageBox.Show("File name must be between 2 and 25 characters in length.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public void ExportPDF(DataTable dtTable, String strPdfPath, String strHeader)
        {
            System.IO.FileStream fs = new FileStream(strPdfPath, FileMode.Create, FileAccess.Write, FileShare.None);    //Opens a new file stream
            Document document = new Document(); //Initialises new document class
            document.SetPageSize(iTextSharp.text.PageSize.A2);  //Sets page size of the PDF to A2 to accomidate for the amount of rows
            PdfWriter writer = PdfWriter.GetInstance(document, fs); //Gets a new instance of the pdf writer
            document.Open();    //Opens the document

            PdfPTable table = new PdfPTable(dtTable.Columns.Count); //Creates a new table of the amount of columns our currently displayed data table contains
            //Table header
            Paragraph prgHeading = new Paragraph(); //New paragraph of page heading
            prgHeading.Alignment = Element.ALIGN_CENTER;    //Set allignment of heading to center
            prgHeading.Add(new Chunk(strHeader.ToUpper())); 
            document.Add(new Paragraph(" "));
            document.Add(prgHeading);   //Writes the paragraph header

            //Cycles through the data table columns to write the column headings in the file
            for (int i = 0; i < dtTable.Columns.Count; i++)
            {
                //Creates a new cell
                PdfPCell cell = new PdfPCell();
                cell.BackgroundColor = iTextSharp.text.BaseColor.GRAY;  //Sets header column colour to grey to differentiate them
                cell.AddElement(new Chunk(dtTable.Columns[i].ColumnName.ToUpper()));    //Sets columns to upper
                table.AddCell(cell);    //Adds cell to the table
            }
            //Table Data
            for (int i = 0; i < dtTable.Rows.Count; i++)
            {
                for (int j = 0; j < dtTable.Columns.Count; j++)
                {
                    //Cycles through each row and then each column of each row to write to the table
                    table.AddCell(dtTable.Rows[i][j].ToString());
                }
            }

            document.Add(table);    //Adds the table to the file
            document.Close();   //Closes the document
            writer.Close(); //Closes the writer
            fs.Close(); //Closes the file stream
            MessageBox.Show("Finished writing to PDF file!", "PDF Exporter", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Process.Start(strPdfPath);  //Opens PDF file at end of process
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Opens folder dialouge to export pdf to
            DialogResult result = folderPath.ShowDialog();

            directoryPath = folderPath.SelectedPath;

            label2.Text = directoryPath;    //Shows directory selected
        }

        //On click to check if name is valid
        private void button2_Click(object sender, EventArgs e)
        {
            CheckValidPDFFileName();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //Assigns pdf name as input is changed
            fileNamePDF = textBox1.Text;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = dataTableBackUp; //Sets the data source to the back up we created when the original file is opened 
            sqlStatement = "";  //Resets input to empty
            sqlStatement = Interaction.InputBox("Enter an SQL command", "SQL input", "SELECT * FROM " + dataTableMB.Rows[0][2].ToString()); //Default SQL command is select all from the current table of opened file

            dbSource = "Data source = " + directoryPathMDB; //Gets path and datasource of current file to query
            dbProvider = "Provider=Microsoft.Jet.OLEDB.4.0;";   //Gets provider
            dbMBCon.ConnectionString = dbProvider + dbSource;   //Establishes connection string
            dbMBCon.Open(); //Opens connection
            try
            {
                ((DataTable)dataGridView1.DataSource).Rows.Clear(); //Clears the current datasource
                OleDbDataAdapter dt = new OleDbDataAdapter(sqlStatement, dbMBCon);  //Creates new data adapter to write read results to data grid view
                dt.Fill(dataTable); //Data table filled
                dataGridView1.DataSource = dataTable;   //Data table set to source
                dbMBCon.Close();    //Connection to file closed
                dataGridView1.Refresh();    //Data grid refreshed
            }
            //Catches exception of SQL statement being invalid
            catch (System.Data.OleDb.OleDbException)
            {
                dbMBCon.Close();    //File is closed
                MessageBox.Show("SQL statement not recognised." ,"Invalid entry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
