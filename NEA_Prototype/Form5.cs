using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Xml;
using System.Windows.Forms;
using System.IO;
using ADOX;
using System.Data.OleDb;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Threading;
using System.Diagnostics;


namespace NEA_Prototype
{

    public partial class Form5 : Form
    {
        DataTable savDataTable = new DataTable();   //Datatable to be written to and displayed
        FolderBrowserDialog fPath = new FolderBrowserDialog();  //Allows the user to select a file location to save to
        Report displayReport = new Report();    //New instance of report class 

        Thread th;  //Thread to take back to main menu
        private bool validSaveName = false; //Bool to store whether a valid save name was enterered

        string fileName;    //Stores the user inputted name for the file to save to
        string directoryPath;   //Stores the user inputted path 

        OleDbConnection dbCon = new System.Data.OleDb.OleDbConnection();    //New connection to sql file
        OleDbConnection dbMBCon = new System.Data.OleDb.OleDbConnection();
        string dbProvider;  //Microsoft provider
        string dbSource;    //Specifies the data source for the file
        string db;  //Stores the name of the database
        string dbMB;    //Stores the name of the master database
        string dbPath;  //Stores the full path of the database
        private List<String> illegalCharacters = new List<string>() { "<", ">", ":", Char.ConvertFromUtf32(34).ToString(), @"\", @"/", "|", "?", "*" };
        DateTime curDate;
        private string isPriv;  //Variable for the check box to denote whether the report we are writing is only accesible by the user
        string MDBConString;  //Compiled full connection string


        public CreateUser user = new CreateUser();


        public Form5()
        {
            InitializeComponent();
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        //Method that takes in the report class and displays the datatable stored in the instance
        public void DisplayReport(Report report) {  
            displayReport = report;
            dataGridView1.DataSource = report.dataTable;
            savDataTable = report.dataTable;
        }

        private void Form5_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        //Method that stores the result of the user choosing a file
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = fPath.ShowDialog();

            directoryPath = fPath.SelectedPath;

            label1.Text = directoryPath;
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e) {

        }

        //Displays the file name on the form
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            fileName = textBox1.Text;
        }

        //Method to check whether the report name is valid
        private void CheckValidReportFileName()
        {   
            //Checks the length of the file name is between 2 and 25 characters long
            if(textBox1.Text.Length >= 2 & textBox1.Text.Length <= 25)
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
                //If the save name is valid and the directory path is valid the mdb is exported
                if (validSaveName == true)
                    if (directoryPath != null)
                        //Save report method is called
                        SaveReportFile();
                    else
                        MessageBox.Show("File location cannot be blank.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
                MessageBox.Show("File name cannot be blank.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                        ExportPDF(savDataTable, (directoryPath + "\\" + fileName + ".pdf"), ("Report Author: " + user.e_firstName + " " + user.e_lastName + " " + "Department: " + user.e_role));
                    else
                        MessageBox.Show("File location cannot be blank.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
                MessageBox.Show("File name cannot be blank.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        private void SaveReportFile()
        {
            ADOX.CatalogClass tCat = new CatalogClass();    //New instance of a database class

            dbProvider = "Provider=Microsoft.Jet.OLEDB.4.0;";   //The microsoft provider

            //Checks whether the file already exists in the directory, if it does "_clone" is added to the end and the file is created
            try
            {
                db = @"\\" + fileName + ".mdb";
                dbPath = directoryPath + db;
                MDBConString = dbProvider + "Data source = " + dbPath + ";Jet OLEDB:Engine Type=5";  //Compiled full connection string
                tCat.Create(MDBConString);  //Creates a new database file using the connection string
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                db = @"\\" + fileName + "_clone" + ".mdb";
                dbPath = directoryPath + db;
                MDBConString = dbProvider + "Data source = " + dbPath + ";Jet OLEDB:Engine Type=5";  //Compiled full connection string
                tCat.Create(MDBConString);  //Creates a new database file using the connection string
            }
            finally
            {
                dbSource = "Data source = " + dbPath;
                dbCon.ConnectionString = dbProvider + dbSource;

                System.Runtime.InteropServices.Marshal.ReleaseComObject(tCat);
                tCat = null;
                GC.Collect();

                tCat = new ADOX.CatalogClass();
                tCat.let_ActiveConnection(MDBConString);
            }

            //Checks the report code
            if (displayReport.e_reportType == 0)
            {
                //Creates a new table instance and adds the columns
                ADOX.Table tTable = new ADOX.Table();
                tTable.Name = "FORDAYDEM";
                tTable.Columns.Append("National Boundary Identifier", DataTypeEnum.adVarWChar, 10);
                tTable.Columns.Append("Settlement Date", DataTypeEnum.adVarWChar, 10);
                tTable.Columns.Append("Settlement Period", DataTypeEnum.adVarWChar, 10);
                tTable.Columns.Append("Record Type", DataTypeEnum.adVarWChar, 10);
                tTable.Columns.Append("Publishing Period Commencing Time", DataTypeEnum.adVarWChar, 20);
                tTable.Columns.Append("Demand", DataTypeEnum.adVarWChar, 10);
                tTable.Columns.Append("Spn Demand", DataTypeEnum.adVarWChar, 10);
                tTable.Columns.Append("Spn Generation", DataTypeEnum.adVarWChar, 10);
                tTable.Columns.Append("Active Flag", DataTypeEnum.adVarWChar, 3);
                tCat.Tables.Append((object)tTable);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tCat);
                tCat = null;
                GC.Collect();

                int ID = 0;

                OleDbCommand cmd = dbCon.CreateCommand();
                dbCon.Open();

                foreach (XmlNode xn in displayReport.xnList)
                {
                    //Declares variables to write to file
                    ID += 1;
                    string nationalBoundaryIdentifier = xn["nationalBoundaryIdentifier"].InnerText;
                    string settlementDate = xn["settlementDate"].InnerText;
                    string settlementPeriod = xn["settlementPeriod"].InnerText;
                    string recordType = xn["recordType"].InnerText;
                    string publishingPeriodCommencingTime = xn["publishingPeriodCommencingTime"].InnerText;
                    string demand;
                    string spnDemand = "";
                    string spnGeneration = "";
                    try
                    {
                        demand = xn["demand"].InnerText;
                    }
                    catch (System.NullReferenceException)
                    {
                        demand = "NULL";
                    }
                    try
                    {
                        spnDemand = xn["spnDemand"].InnerText;
                    }
                    catch (System.NullReferenceException)
                    {
                        spnDemand = "NULL";
                    }
                    try
                    {
                        spnGeneration = xn["spnGeneration"].InnerText;
                    }
                    catch (System.NullReferenceException)
                    {
                        spnGeneration = "NULL";
                    }
                    string activeFlag = xn["activeFlag"].InnerText;

                    //The SQL command to execute to write to the file
                    cmd.CommandText = "INSERT INTO FORDAYDEM([National Boundary Identifier],[Settlement Date],[Settlement Period],[Record Type],[Publishing Period Commencing Time],[Demand],[Spn Demand],[Spn Generation],[Active Flag])VALUES('" + nationalBoundaryIdentifier + "', '" + settlementDate + "', '" + settlementPeriod + "', '" + recordType + "', '" + publishingPeriodCommencingTime + "', '" + demand + "', '" + spnDemand + "', '" + spnGeneration + "', '" + activeFlag + "')";

                    //Executing the command
                    cmd.Connection = dbCon;
                    cmd.ExecuteNonQuery();
                }
                //Closing the connection
                dbCon.Close();

                //Opens the master database file
                dbMB = @"\\MADB.mdb";
                dbPath = Environment.CurrentDirectory + dbMB;
                dbSource = "Data source = " + dbPath;
                dbMBCon.ConnectionString = dbProvider + dbSource;
                OleDbCommand cmdMB = dbMBCon.CreateCommand();
                //Checks whether the isPriv box is checked
                if (checkBox1.Checked == true)
                    isPriv = "true";
                else
                    isPriv = "false";
                //Gets current time
                curDate = DateTime.Now;
                //Command string to insert into database and executed
                cmdMB.CommandText = "INSERT INTO FILERECORDS([File Name],[User Name],[Report Type],[Full Name],[Role],[Date Created],[Private])VALUES('" + fileName + "', '" + user.e_username + "', '" + "FORDAYDEM" + "', '" + user.e_firstName + user.e_lastName + "', '" + user.e_role + "', '" + curDate + "', '" + isPriv + "')";
                cmdMB.CommandType = CommandType.Text;
                dbMBCon.Open();

                //Executing the command
                cmdMB.Connection = dbMBCon;
                cmdMB.ExecuteNonQuery();
                dbMBCon.Close();
                MessageBox.Show("Results written to file!", "File exporter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //Opens file
                Process.Start(dbPath);
            }
            else if (displayReport.e_reportType == 1)
            {
                //Creates a new table instance and adds the columns
                ADOX.Table tTable = new ADOX.Table();
                tTable.Name = "B1440";
                tTable.Columns.Append("Time Series ID", DataTypeEnum.adVarWChar, 40);
                tTable.Columns.Append("Business Type", DataTypeEnum.adVarWChar, 20);
                tTable.Columns.Append("Power System Resource Type", DataTypeEnum.adVarWChar, 20);
                tTable.Columns.Append("Settlement Date", DataTypeEnum.adVarWChar, 20);
                tTable.Columns.Append("Process Type", DataTypeEnum.adVarWChar, 20);
                tTable.Columns.Append("Settlement Period", DataTypeEnum.adVarWChar, 10);
                tTable.Columns.Append("Quantity", DataTypeEnum.adVarWChar, 10);
                tTable.Columns.Append("Document Type", DataTypeEnum.adVarWChar, 40);
                tTable.Columns.Append("Curve Type", DataTypeEnum.adVarWChar, 40);
                tTable.Columns.Append("Resolution", DataTypeEnum.adVarWChar, 20);
                tTable.Columns.Append("Active Flag", DataTypeEnum.adVarWChar, 20);
                tTable.Columns.Append("Document ID", DataTypeEnum.adVarWChar, 40);
                tTable.Columns.Append("Document Review Number", DataTypeEnum.adVarWChar, 10);
                tCat.Tables.Append((object)tTable);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tCat);
                tCat = null;
                GC.Collect();

                OleDbCommand cmd = dbCon.CreateCommand();
                dbCon.Open();

                foreach (XmlNode xn in displayReport.xnList)
                {
                    //Declares variables to write to file
                    string timeSeriesID = xn["timeSeriesID"].InnerText;
                    string businessType = xn["businessType"].InnerText;
                    string powerSystemResourceType = xn["powerSystemResourceType"].InnerText;
                    string settlementDate = xn["settlementDate"].InnerText;
                    string processType = xn["processType"].InnerText;
                    string settlementPeriod = xn["settlementPeriod"].InnerText;
                    string quantity = xn["quantity"].InnerText;
                    string documentType = xn["documentType"].InnerText;
                    string curveType = xn["curveType"].InnerText;
                    string resolution = xn["resolution"].InnerText;
                    string activeFlag = xn["activeFlag"].InnerText;
                    string documentID = xn["documentID"].InnerText;
                    string documentRevNum = xn["documentRevNum"].InnerText;

                    //The SQL command to execute to write to the file
                    cmd.CommandText = "INSERT INTO B1440([Time Series ID],[Business Type],[Power System Resource Type],[Settlement Date],[Process Type],[Settlement Period],[Quantity],[Document Type],[Curve Type],[Resolution],[Active Flag],[Document ID],[Document Review Number])VALUES('" + timeSeriesID + "', '" + businessType + "', '" + powerSystemResourceType + "', '" + settlementDate + "', '" + processType + "', '" + settlementPeriod + "', '" + quantity + "', '" + documentType + "','" + curveType + "','" + resolution + "','" + activeFlag + "','" + documentID + "','" + documentRevNum + "')";

                    //Executing the command
                    cmd.Connection = dbCon;
                    cmd.ExecuteNonQuery();
                }
                //Closing the connection
                dbCon.Close();

                dbMB = @"\\MADB.mdb";
                dbPath = Environment.CurrentDirectory + dbMB;
                dbSource = "Data source = " + dbPath;
                dbMBCon.ConnectionString = dbProvider + dbSource;
                OleDbCommand cmdMB = dbMBCon.CreateCommand();
                if (checkBox1.Checked == true)
                    isPriv = "true";
                else
                    isPriv = "false";
                //Gets current time
                curDate = DateTime.Now;
                //Command string to insert into database and executed
                cmdMB.CommandText = "INSERT INTO FILERECORDS([File Name],[User Name],[Report Type],[Full Name],[Role],[Date Created],[Private])VALUES('" + fileName + "', '" + user.e_username + "', '" + "B1440" + "', '" + user.e_firstName + user.e_lastName + "', '" + user.e_role + "', '" + curDate + "', '" + isPriv + "')";
                cmdMB.CommandType = CommandType.Text;
                dbMBCon.Open();

                //Executing the command
                cmdMB.Connection = dbMBCon;
                cmdMB.ExecuteNonQuery();
                dbMBCon.Close();
                MessageBox.Show("Results written to file!", "File exporter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //Opens file
                Process.Start(dbPath);
            }
            else if (displayReport.e_reportType == 2)
            {
                //Creates a new table instance and adds the columns
                ADOX.Table tTable = new ADOX.Table();
                tTable.Name = "FOU2T52W";
                tTable.Columns.Append("Record", DataTypeEnum.adVarWChar, 10);
                tTable.Columns.Append("Fuel Type", DataTypeEnum.adVarWChar, 10);
                tTable.Columns.Append("Publishing Period Commencing Time", DataTypeEnum.adVarWChar, 20);
                tTable.Columns.Append("System Zone", DataTypeEnum.adVarWChar, 3);
                tTable.Columns.Append("Calandar Week Number", DataTypeEnum.adVarWChar, 3);
                tTable.Columns.Append("Year", DataTypeEnum.adVarWChar, 4);
                tTable.Columns.Append("Output Usable", DataTypeEnum.adVarWChar, 8);
                tTable.Columns.Append("Active Flag", DataTypeEnum.adVarWChar, 3);
                tCat.Tables.Append((object)tTable);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tCat);
                tCat = null;
                GC.Collect();

                OleDbCommand cmd = dbCon.CreateCommand();
                dbCon.Open();

                foreach (XmlNode xn in displayReport.xnList)
                {
                    //Declares variables to write to file
                    string recordType = xn["recordType"].InnerText;
                    string fuelType = xn["fuelType"].InnerText;
                    string publishingPeriodCommencingTime = xn["publishingPeriodCommencingTime"].InnerText;
                    string systemZone = xn["systemZone"].InnerText;
                    string calendarWeekNumber = xn["calendarWeekNumber"].InnerText;
                    string year = xn["year"].InnerText;
                    string outputUsable = xn["outputUsable"].InnerText;
                    string activeFlag = xn["activeFlag"].InnerText;

                    //The SQL command to execute to write to the file
                    cmd.CommandText = "INSERT INTO FOU2T52W([Record],[Fuel Type],[Publishing Period Commencing Time],[System Zone],[Calandar Week Number],[Year],[Output Usable],[Active Flag])VALUES('" + recordType + "', '" + fuelType + "', '" + publishingPeriodCommencingTime + "', '" + systemZone + "', '" + calendarWeekNumber + "', '" + year + "', '" + outputUsable + "', '" + activeFlag + "')";

                    //Executing the command
                    cmd.Connection = dbCon;
                    cmd.ExecuteNonQuery();
                }
                //Closing the connection
                dbCon.Close();

                dbMB = @"\\MADB.mdb";
                dbPath = Environment.CurrentDirectory + dbMB;
                dbSource = "Data source = " + dbPath;
                dbMBCon.ConnectionString = dbProvider + dbSource;
                OleDbCommand cmdMB = dbMBCon.CreateCommand();
                if (checkBox1.Checked == true)
                    isPriv = "true";
                else
                    isPriv = "false";
                //Gets current date and time
                curDate = DateTime.Now;
                //Command string to insert into database and executed
                cmdMB.CommandText = "INSERT INTO FILERECORDS([File Name],[User Name],[Report Type],[Full Name],[Role],[Date Created],[Private])VALUES('" + fileName + "', '" + user.e_username + "', '" + "FOU2T52W" + "', '" + user.e_firstName + user.e_lastName + "', '" + user.e_role + "', '" + curDate + "', '" + isPriv + "')";
                cmdMB.CommandType = CommandType.Text;
                dbMBCon.Open();

                cmdMB.Connection = dbMBCon;
                cmdMB.ExecuteNonQuery();
                dbMBCon.Close();
                MessageBox.Show("Results written to file!", "File exporter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //Opens file
                Process.Start(dbPath);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CheckValidReportFileName();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Opens main menu
            this.Close();
            th = new Thread(openForm3);
            th.SetApartmentState(ApartmentState.STA);
            th.Start();
        }

        private void openForm3(object obj)
        {
            Application.Run(new Form3());
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        //Checks the valid PDF name 
        private void button4_Click(object sender, EventArgs e)
        {
            CheckValidPDFFileName();
        }
    }
}
