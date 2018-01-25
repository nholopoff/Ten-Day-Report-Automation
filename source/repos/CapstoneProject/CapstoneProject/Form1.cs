
/* TODO
 * DataSource to pull in information from spreadsheet
 * Professor selection handled by listview
 *  LAST NAME, FIRST NAME, ...
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
                Filter = "Microsoft Excel 2007-2013 XML (*.xlsx)|*.xlsx|All Files (*.*)|*.*", //update this as file types are supported
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
                                                  "Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\""; //connectionstrings.com


                        OleDbConnection connection = new OleDbConnection(connectionString); // establish link to file
                        OleDbDataAdapter command = new OleDbDataAdapter("select * from [Sheet1$]", connection); // select every cell on Sheet1
                        DataSet ds = new DataSet(); // initialize dataset for data grid
                        command.Fill(ds); // fill data grid
                        scheduleDataGridView.DataSource = ds.Tables[0];
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
    }
}

