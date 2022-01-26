using System.Text;
using System.Net;
using System.Xml;
using System.IO;
using System.Data.OleDb;
using System.Data;
using System.Diagnostics;


namespace NEA_Prototype
{
    public class Report
    {
        public int e_reportType;     //Report type represented in one of three numbers
        public string e_period;     //How far back does the data go
        private string reportCode;  //Report type string represented on BRMS database 
        public string e_settlementDate;     //The date to search from +/- period
        public string e_dbFile;     //Name of the database file to write to
        OleDbConnection dbCon = new OleDbConnection();      //API used to connect to database
        public DataTable dataTable = new DataTable();   //The data table that will be filled from scraped results from web database to be displayed
        public XmlNodeList xnList;  //Used to store the retrived results of the web database which is stored via XML     

        public void GenerateReport(int reportType, string settlement, string period)
        {
            //Assigning variables used to query web database
            e_reportType = reportType;
            e_settlementDate = settlement;
            e_period = period;

            //Checking the report type selected from the UI menu and finding the matching report code
            if (reportType == 0)
            {
                reportCode = "FORDAYDEM";
            }

            else if (reportType == 1)
            {
                reportCode = "B1440";
            }
            else if (reportType == 2)
            {
                reportCode = "FOU2T52W";
            }

            //Subbing into the base url the parameters to query the web database
            string e_queryUrl = $"https://api.bmreports.com/BMRS/{reportCode}/V1?APIKey=r90kuflzcdja4gu&SettlementDate={e_settlementDate}&Period={e_period}&ServiveType=xml/XML/csv/CSV";

            var dataReq = (HttpWebRequest)WebRequest.Create(e_queryUrl);  //Assigning the webrequest with the database access url
            dataReq.ContentType = "CSV";
            dataReq.Method = "GET";   //Assigning what we want to do with the webrequest

            HttpWebResponse dataResp = (HttpWebResponse)dataReq.GetResponse();  //Returning the response from the webrequest
            Stream receiveStream = dataResp.GetResponseStream();    //Opening a stream to read the contents of the webrepsonse 
            StreamReader reader = new StreamReader(receiveStream, Encoding.UTF8);   //Assigning a reader to read the stream with unicode encoding for the characters
            string e_response = reader.ReadToEnd(); //String that will contain the raw contents of the webpage that the reader read
            string e_parsed = null;   //String that will store the trimmed contents of the webpage after the header is removed 
            int indexOfResponse = e_response.IndexOf("<responseList>") - 14;    //Gets the index of every item in the webpage before "<responseList>"

            //Checks if we successfully connected to the database

            if (dataResp.StatusCode.ToString() == "OK")
            {
                //Checks if we are already at the start of the xml data from the database
                if (indexOfResponse >= 0)
                {
                    //Trims the header of the page
                    indexOfResponse += "<responseList>".Length;
                    int indexOfBodyEnd = e_response.IndexOf("</responseList>") + 15;
                    if (indexOfBodyEnd >= 0)
                        e_parsed = e_response.Substring(indexOfResponse, indexOfBodyEnd - indexOfResponse);
                }
                else
                {
                    e_parsed = e_response.Substring(indexOfResponse);
                }

                XmlDocument e_report = new XmlDocument();   //Creates a new xml document which will store our webpages result
                e_report.LoadXml(e_parsed); //Loads the contents of e_parsed into the xml doc

                xnList = e_report.SelectNodes("responseList/item"); //Create a list of all the nodes we need in this case the ones called "item"

                //check what report type we are going to display as each has different attributes
                if (reportType == 0)
                {
                    //Assign data columns with corresponding values in webpage database and add then to our datatable
                    DataColumn[] columns = { new DataColumn("National Boundary Identifier"), new DataColumn("Settlement Date"), new DataColumn("Settlement Period"), new DataColumn("Record Type"), new DataColumn("Publishing Period Commencing Time"), new DataColumn("Demand"), new DataColumn("Spn Demand"), new DataColumn("Spn Generation"), new DataColumn("activeFlag") };
                    dataTable.Columns.AddRange(columns);

                    //Cycle through each xml node in our list
                    foreach (XmlNode xn in xnList)
                    {
                        //Declare and assign strings to corresponding xml nodes
                        string nationalBoundaryIdentifier = xn["nationalBoundaryIdentifier"].InnerText;
                        string settlementDate = xn["settlementDate"].InnerText;
                        string settlementPeriod = xn["settlementPeriod"].InnerText;
                        string recordType = xn["recordType"].InnerText;
                        string publishingPeriodCommencingTime = xn["publishingPeriodCommencingTime"].InnerText;
                        string demand;
                        string spnDemand = "";
                        string spnGeneration = "";
                        //some entities in this database dont share the same attributes so this checks the right attribute to be written
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

                        //Add row to table
                        object[] row = { nationalBoundaryIdentifier, settlementDate, settlementPeriod, recordType, publishingPeriodCommencingTime, demand, spnDemand, spnGeneration, activeFlag };
                        dataTable.Rows.Add(row);

                    }
                }
                else if (reportType == 1)
                {
                    //Assign data columns with corresponding values in webpage database and add then to our datatable
                    DataColumn[] columns = { new DataColumn("Time Series ID"), new DataColumn("Business Type"), new DataColumn("Power System Resource Type"), new DataColumn("Settlement Date"), new DataColumn("Process Type"), new DataColumn("Settlement Period"), new DataColumn("Quantity"), new DataColumn("Document Type"), new DataColumn("Curve Type"), new DataColumn("Resolution"), new DataColumn("Active Flag"), new DataColumn("Document ID"), new DataColumn("Document Rev Num") };
                    dataTable.Columns.AddRange(columns);

                    //Cycle through each xml node in our list

                    foreach (XmlNode xn in xnList)
                    {
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

                        //Add row to table
                        object[] row = { timeSeriesID, businessType, powerSystemResourceType, settlementDate, processType, settlementPeriod, quantity, documentType, curveType, resolution, activeFlag, documentID, documentRevNum };
                        dataTable.Rows.Add(row);
                    }
                }
                else if (reportType == 2)
                {
                    //Assign data columns with corresponding values in webpage database and add then to our datatable

                    DataColumn[] columns = { new DataColumn("Record"), new DataColumn("Fuel Type"), new DataColumn("Publishing Period Commencing Time"), new DataColumn("System Zone"), new DataColumn("Calendar Week Number"), new DataColumn("Year"), new DataColumn("Output Usable"), new DataColumn("Active Flag") };
                    dataTable.Columns.AddRange(columns);

                    //Cycle through each xml node in our list

                    foreach (XmlNode xn in xnList)
                    {
                        string recordType = xn["recordType"].InnerText;
                        string fuelType = xn["fuelType"].InnerText;
                        string publishingPeriodCommencingTime = xn["publishingPeriodCommencingTime"].InnerText;
                        string systemZone = xn["systemZone"].InnerText;
                        string calendarWeekNumber = xn["calendarWeekNumber"].InnerText;
                        string year = xn["year"].InnerText;
                        string outputUsable = xn["outputUsable"].InnerText;
                        string activeFlag = xn["activeFlag"].InnerText;

                        //Add row to table
                        object[] row = { recordType, fuelType, publishingPeriodCommencingTime, systemZone, calendarWeekNumber, year, outputUsable, activeFlag };
                        dataTable.Rows.Add(row);
                    }
                }
            }
        }
    }
}
