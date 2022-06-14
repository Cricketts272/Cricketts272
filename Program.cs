using IronXL;
using System.Data;
using System;
using System.Text;

// input path, volume, and output folder path (must be a .xlsx file)
string path = "/Users/clairricketts/Documents/Picklist_generator_input/input_test_96plex.xlsx";
int plex = 96;
string volume = "45";
string outPathFolder = "/Users/clairricketts/Documents/Picklist_generator_output/";
// instantiating a dictionary to hold the sequences
Dictionary<string, string> sequence = new Dictionary<string, string>();
// loading the .xlsx file
WorkBook wb = WorkBook.Load(path); // Your Excel file Name
WorkSheet ws = wb.GetWorkSheet("Sheet1");
DataTable dt = ws.ToDataTable(true);//parse sheet1 of sample.xlsx file into datatable


if (plex == 8)
{



    for (int i = 0; i <= dt.Rows.Count - 1; i++) // Put the data table values into an iterable dictionary type
    {
        string key = String.Format("seq{0}", i.ToString());
        sequence.Add(key, dt.Rows[i]["sequence"].ToString());


    }

    Dictionary<string, char[]> seqarr = new Dictionary<string, char[]>();
    for (int i = 0; i <= dt.Rows.Count - 1; i++) // breaking up each dictionary entry into a char array
    {
        string keys = String.Format("seq{0}", i.ToString());
        string key = String.Format("{0}", i.ToString());
        seqarr.Add(key, sequence[keys].ToCharArray());
    }


    // determine the number of plates (picklists) we need to generate
    int numberOfPlates = seqarr["0"].Length / 48;



    // creates a data table with the appropriate headers
    for (int q = 0; q <= numberOfPlates; q++)
    {   // arrcount variable provides a lower bound when sifting through sequence arrays
        int arrcount = q * 48;


        DataTable pickList = new DataTable();
        pickList.Clear();
        pickList.Columns.Add("Well");
        pickList.Columns.Add("Volume");
        pickList.Columns.Add("Reagent");
        // creating a string that contains row names
        string[] plateRow = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P" };

        DataRow row;
        //
        for (int j = 0; j <= (dt.Rows.Count) - 1; j++)
        {
            string count = j.ToString();
            int bound = (seqarr[count].Length);

            int upperBound = 0;
            if (bound <= ((q + 1) * 48))
            {
                upperBound = bound;
            }
            if (bound > ((q + 1) * 48))
            {
                upperBound = ((q + 1) * 48);
            }



            for (int i = (1 + arrcount); i < upperBound; i++)
            {
                int g = i / 24;
                int multiple = 1; // ensures we get all odd multiples of 24 on the picklist
                if (q % 2 != 0)
                {
                    multiple = q + 2;
                }
                Console.WriteLine(multiple);
                if (g % 2 == 0 || i == 24 * multiple)
                {
                    row = pickList.NewRow();
                    row["Reagent"] = seqarr[count][i - 1];
                    row["Volume"] = volume;
                    row["Well"] = String.Format("{0}{1}", plateRow[j], (i - arrcount).ToString());
                    pickList.Rows.Add(row);
                }

                if (g % 2 != 0)
                {
                    int p = (i - arrcount + 1) - 24;
                    row = pickList.NewRow();
                    row["Reagent"] = seqarr[count][i - 1];
                    row["Volume"] = volume;
                    row["Well"] = String.Format("{0}{1}", plateRow[j + 8], p.ToString());
                    pickList.Rows.Add(row);
                }

            }
        }







        //  Write to .csv file
        string outPathFile = String.Format("Tempest_picklist_{0}.csv", q.ToString());
        Console.WriteLine(outPathFile);

        StringBuilder sb = new StringBuilder();

        IEnumerable<string> columnNames = pickList.Columns.Cast<DataColumn>().
                                          Select(column => column.ColumnName);
        sb.AppendLine(string.Join(",", columnNames));

        foreach (DataRow cell in pickList.Rows)
        {
            IEnumerable<string> fields = cell.ItemArray.Select(field => field.ToString());
            sb.AppendLine(string.Join(",", fields));
        }
        // .csv output path
        File.WriteAllText(outPathFolder + outPathFile, sb.ToString());

    }
}


if (plex == 16)
{



    for (int i = 0; i <= dt.Rows.Count - 1; i++) // Put the data table values into an iterable dictionary type
    {
        string key = String.Format("seq{0}", i.ToString());
        sequence.Add(key, dt.Rows[i]["sequence"].ToString());


    }

    Dictionary<string, char[]> seqarr = new Dictionary<string, char[]>();
    for (int i = 0; i <= dt.Rows.Count - 1; i++) // breaking up each dictionary entry into a char array
    {
        string keys = String.Format("seq{0}", i.ToString());
        string key = String.Format("{0}", i.ToString());
        seqarr.Add(key, sequence[keys].ToCharArray());
    }


    // determine the number of plates (picklists) we need to generate
    int numberOfPlates = seqarr["0"].Length / 24;



    // creates a data table with the appropriate headers
    for (int q = 0; q <= numberOfPlates; q++)
    {   // arrcount variable provides a lower bound when sifting through sequence arrays
        int arrcount = q * 24;


        DataTable pickList = new DataTable();
        pickList.Clear();
        pickList.Columns.Add("Well");
        pickList.Columns.Add("Volume");
        pickList.Columns.Add("Reagent");
        // creating a string that contains row names
        string[] plateRow = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P" };

        DataRow row;
        for (int j = 0; j <= (dt.Rows.Count) - 1; j++)
        {
            string count = j.ToString();
            int bound = (seqarr[count].Length);
            int upperBound = 0;
            // decide the upper bound of the sequence array we are transforming 
            if (bound <= ((q + 1) * 24))
            {
                upperBound = bound;
            }
            if (bound > ((q + 1) * 24))
            {
                upperBound = ((q + 1) * 24);
            }



            
            for (int i = (1 + arrcount); i <= upperBound; i++)
            {

                row = pickList.NewRow();
                row["Reagent"] = seqarr[count][i - 1];
                row["Volume"] = volume;
                row["Well"] = String.Format("{0}{1}", plateRow[j], (i - arrcount).ToString());
                pickList.Rows.Add(row);
               
            }
        }
    







        //  Write to .csv file
        string outPathFile = String.Format("Tempest_picklist_{0}.csv", q.ToString());
        Console.WriteLine(outPathFile);

        StringBuilder sb = new StringBuilder();

        IEnumerable<string> columnNames = pickList.Columns.Cast<DataColumn>().
                                          Select(column => column.ColumnName);
        sb.AppendLine(string.Join(",", columnNames));

        foreach (DataRow cell in pickList.Rows)
        {
            IEnumerable<string> fields = cell.ItemArray.Select(field => field.ToString());
            sb.AppendLine(string.Join(",", fields));
        }
        // .csv output path
        File.WriteAllText(outPathFolder + outPathFile, sb.ToString());

    }
}


if (plex == 96)
{
    string[] plateRow = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P" };

    for (int i = 0; i <= dt.Rows.Count - 1; i++) // Put the data table values into an iterable dictionary type
    {
        string key = String.Format("seq{0}", i.ToString());
        sequence.Add(key, dt.Rows[i]["sequence"].ToString());
    }

    Dictionary<string, char[]> seqarr = new Dictionary<string, char[]>();
    for (int i = 0; i <= dt.Rows.Count - 1; i++) // breaking up each dictionary entry into a char array
    {
        string keys = String.Format("seq{0}", i.ToString());
        string key = String.Format("{0}", i.ToString());
        seqarr.Add(key, sequence[keys].ToCharArray());
    }

    // determine the number of plates (picklists) we need to generate
    int numberOfPlates = seqarr["0"].Length / 4; // ***it assumes that the first sequence is the longest
    // creates a data table with the appropriate headers
    for (int q = 0; q <= numberOfPlates; q++)
    {   // arrcount variable provides a lower bound when sifting through sequence arrays
        int arrcount = q * 4;


        DataTable pickList = new DataTable();
        pickList.Clear();
        pickList.Columns.Add("Well");
        pickList.Columns.Add("Volume");
        pickList.Columns.Add("Reagent");
        List<string> well = new List<string>();
        List<string> reagent = new List<string>();
        // due to the nature of 96 plex, we need to rearrage the sequence arrays
        //and line them up with the proper well label
        // the entire 384 well plate is split up into 4 sections of 6X16 wells
        // where each section represents one extension for each sequence (or each rod)
        // for example, the first 16 sequences would have their first 4 extensions in columns 1,7,13,19
        // the next 16 sequences would have their first 4 extensions in columns 2,8,14,20
        // then 3,9,15,21 --> 4,10,16,22 --> 5,11,17,23 --> 6,12,18,24
       
        for (int m = 0; m <= (dt.Rows.Count) - 1; m++)
        {
            string count = m.ToString();
            int bound = (seqarr[count].Length);
            int upperBound = 0;
            // decide the upper bound of the sequence array we are transforming 
            if (bound <= ((q + 1) * 4))
            {
                upperBound = bound;
            }
            if (bound > ((q + 1) * 4))
            {
                upperBound = ((q + 1) * 4);
            }
           
            for (int y = (1 + arrcount); y < upperBound + 1; y++)
            {

                char seqr = seqarr[count][y - 1];
               
                reagent.Add(String.Format("{0}", seqarr[count][y - 1]));
                
            }
            
        }


        for (int k = 0; k < 16; k++) // choosing the well letter
        {
            for (int i = 1; i <= 6; i++) // choosing the lower bound of j
            {
                for (int j = i; j <= i + 19; j += 6) // filling every 6th well 
                {
                    
                    well.Add(String.Format("{0}{1}", plateRow[k], j));
                  
                }

            }



        }

        
        DataRow row;
        for (int i = 0; i < reagent.Count; i++)
        {
            row = pickList.NewRow();
            row["Reagent"] = reagent[i];
            row["Well"] = well[i];
            row["Volume"] = volume;
            pickList.Rows.Add(row);
        }



        

        //  Write to .csv file
        string outPathFile = String.Format("Tempest_picklist_{0}.csv", q.ToString());
        Console.WriteLine(outPathFile);

        StringBuilder sb = new StringBuilder();

        IEnumerable<string> columnNames = pickList.Columns.Cast<DataColumn>().
                                          Select(column => column.ColumnName);
        sb.AppendLine(string.Join(",", columnNames));

        foreach (DataRow cell in pickList.Rows)
        {
            IEnumerable<string> fields = cell.ItemArray.Select(field => field.ToString());
            sb.AppendLine(string.Join(",", fields));
        }
        // .csv output path
        File.WriteAllText(outPathFolder + outPathFile, sb.ToString());

    }
}



