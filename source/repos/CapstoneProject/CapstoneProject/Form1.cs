
/* TODO
 * DataSource to pull in information from spreadsheet
 * Professor selection handled by listview
 * Course output handled by either ListView or DataGrid depending on how much information is needed
 * agree on a design template
 * form goals of application
 * when selected, output the classes they teach, otherwise display whitespace
 * Make buttons work, EDIT ADD REMOVE
 *  add failsafes in case user accidently removes wrong professor
 * Esc throws "are you sure you want to exit? Yes or no"
 * We can populate a DataGridView control with Excel spreadsheet data
 *  OpenFileDialog needed to select spreadsheet data to view
 * Add print capabilities
 * Add compability with other file types
 * CONSIDER PORTING TO EPPlus library instead of OLEDB
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CapstoneProject
{
    public partial class Form1 : Form
    {
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();

        public Form1()
        {
            InitializeComponent();
        }

        /* Esc closes app */
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        /* Prompts user to select database file */
        private void OpenButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdial = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer),
                Title = "Course schedule",
                Filter = "Microsoft Excel 2007-2013 XML (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                FilterIndex = 2,
                RestoreDirectory = true
            };
            
            // Fill DataTable with .xlsb file
            if (fdial.ShowDialog() == DialogResult.OK)
            {
                if (fdial.FileName != null)
                {
                    try
                    {
                        string filename = fdial.FileName.ToString();
                        string connectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + filename + ";" + 
                                                  "Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"";

                        OleDbConnection connection = new OleDbConnection(connectionString);
                        connection.Open();
                        dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                        string sheetName = dt.Rows[0]["TABLE_NAME"].ToString();

                        /* foreach (DataRow dr in dt.Rows)
                        {
                            string query = "SELECT * FROM [" + dr.Item(0).ToString + "]";
                            ds.Clear();
                            OleDbDataAdapter data = new OleDbDataAdapter(query, connection);
                            data.Fill(ds);

                        } */

                        OleDbDataAdapter command = new OleDbDataAdapter(String.Format("SELECT * FROM [{0}]", sheetName), connection);
                        command.Fill(dt);

                        // remove schema information
                        for (int i = 0; i < 12; i++) dt.Columns.Remove(dt.Columns[i]);
                        dt.Rows[0].Delete();
                        dt.AcceptChanges();

                        scheduleDataGridView.DataSource = dt;
                        connection.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        /* Generate Ten Day Report */
        private void GenerateButton_Click(object sender, EventArgs e)
        {
            DataTable outDT = new DataTable();
            outDT.Clear();
            outDT.Columns.Add("Instructor", typeof(String));
            outDT.Columns.Add("Course", typeof(int));
            outDT.Columns.Add("Title", typeof(String));

            DataRow[] row = dt.Select();
            //FIXME
            for (int i = 0; i < row.Length; i++)
            {
                outDT.Rows.Add(dt.Rows[i][i]);
                outDT.Rows.Add(dt.Rows[i][i]);
                outDT.Rows.Add(dt.Rows[i][i]);
            }

            scheduleDataGridView.DataSource = outDT;

            GenReport(outDT, "C:\\Users\\Nick\\Documents\\HERE!\\Output\\report.csv");
        }

        private static void GenReport(DataTable dt, string filePath)
        {
            StreamWriter sw = new StreamWriter(filePath, false);
            //headers  
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sw.Write(dt.Columns[i]);
                if (i < dt.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i])) // makes sure isn't null
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(',')) // allows for ',' characters in cells
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dt.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
    }
}

