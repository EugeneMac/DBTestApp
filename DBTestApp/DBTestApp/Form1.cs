using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Linq;
using System.Data.Linq.Mapping;
using ADOX;
using System.IO;
using System.Data.OleDb;


namespace DBTestApp
{

    public partial class Form1 : Form
    {

        string connectionString = null;
        static private int _activeRow;
        static private string _id;
        static private string _name;
        static private string _description;
        static private string _link;
        static private int[] _checkbox_array;
        static public bool _save;

        BindingSource bs;
        DataTable dtab;
        OleDbDataAdapter dda;
        OleDbCommandBuilder cb;
               
  
        public Form1()
        {
            InitializeComponent();
            connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + Directory.GetCurrentDirectory() + @"\TestDB.mdb";
        }

      
        private void Form1_Load(object sender, EventArgs e)
        {
            //
            // Checking if database Test.mdb file exist
            //
            if (!File.Exists(Directory.GetCurrentDirectory() + @"\TestDB.mdb"))
            {
                //
                // If the database doesn't exist we offer to create one
                //
                if (MessageBox.Show("База данных TestDB не найдена! Создать новую базу?",
                    "ВНИМАНИЕ", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    //
                    // Create DB, Table and all necessary fields
                    //
                    CreateAccessDatabase("TestDB");
                    using (OleDbConnection dbConnect = new OleDbConnection(connectionString))
                    {
                        //
                        // PRIMARY KEY is a MUST! Otherwise no automatic updates from dataGridView could be taken!
                        // AUTOINCREMENT is USEFUL too coz we don't need to keep track of last ID on dataGridView!
                        //
                        OleDbCommand command = dbConnect.CreateCommand();
                        command.CommandText = "create table data (Id AUTOINCREMENT PRIMARY KEY , name MEMO, description MEMO, link MEMO, categories CHAR(50))";
                        try
                        {
                            dbConnect.Open();
                            command.ExecuteNonQuery();
                            dataGridView1.DataSource = connectionString;
                            dataGridView1.Refresh();
                            dbConnect.Close();
                        }
                        
                        catch (Exception exc)
                        {
                            MessageBox.Show(exc.Message);

                        }
                                                                    
                    }
                    
                }
                
            }
            ShowData();
            dataGridView1.DataSource = bs;
        }

        /// <summary>
        /// Showing data from MS Access database
        /// </summary>
        private void ShowData()
        {
            
             try
             {
                    dda = new OleDbDataAdapter("Select * from data", connectionString);
                    cb = new OleDbCommandBuilder(dda);
                    dtab = new DataTable();
                    dda.Fill(dtab);
                    bs = new BindingSource();
                    bs.DataSource = dtab;
                    dataGridView1.DataSource = bs;
              }
              catch (Exception e)
              {
                  MessageBox.Show(e.Message);
              }
        }
    

        /// <summary>
        /// Creating mdb file
        /// </summary>
        /// <param name="filename">the name of file</param>
        private static void CreateAccessDatabase(string filename)
        {
           
            string path = Directory.GetCurrentDirectory();
            if (!Directory.Exists(Directory.GetCurrentDirectory())) Directory.CreateDirectory(path);

            ADOX.Catalog database = new ADOX.Catalog();

            try
            {
                database.Create("Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + path + @"\" + filename + ".mdb; Jet OLEDB:Engine Type=5");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            database = null;
           
        }
        /// <summary>
        /// User clicked on some row in dataGridView1
        /// </summary>
        /// <param name="sender">dataGridView1</param>
        /// <param name="e">Event args</param>
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //
            //static field _save is used to define if the button Сохранить has been pressed
            //
            _save = false;
            frmEditWindow frm = new frmEditWindow();

            //
            //static field _activeRow is used to keep current active row user has pressed
            //
            _activeRow = e.RowIndex;

            //
            //frm PROPERTIES are used to fill all the controls of the frmEditWindow which is form to edit data
            //
            frm.ID = dataGridView1.Rows[_activeRow].Cells[0].Value.ToString();
            frm.NAME = dataGridView1.Rows[_activeRow].Cells[1].Value.ToString();
            frm.DESCRIPTION = dataGridView1.Rows[_activeRow].Cells[2].Value.ToString();
            frm.LINK = dataGridView1.Rows[_activeRow].Cells[3].Value.ToString();

            int[] checkboxes = null;
            //
            // Checking if the CATEGORIES cell is not empty and splitting the data from it with comma, then converting to array
            //
            if (!String.IsNullOrWhiteSpace(dataGridView1.Rows[_activeRow].Cells[4].Value.ToString()))
               checkboxes = (dataGridView1.Rows[_activeRow].Cells[4].Value.ToString()).Split(',').Select(n => Convert.ToInt32(n)).ToArray();
            frm.CHECKBOX_ARRAY = checkboxes;

            frm.ShowDialog();

            // 
            // When the frmEditWindow has been shut saving all data from it to static fields
            //
            _id = frm.ID;
            _name = frm.NAME;
            _description = frm.DESCRIPTION;
            _link = frm.LINK;
            _checkbox_array = frm.CHECKBOX_ARRAY;

            //
            // if Сохранить button has been pressed calling Save procedure
            //
            if (_save) Save();

        }

      
        /// <summary>
        /// Procedure for saving all data collected from frmEditWindow
        /// </summary>

        private void Save()
        {
            //
            // Saving collected data to dataGridView1
            //
            dataGridView1.Rows[_activeRow].Cells[0].Value = _id;
            dataGridView1.Rows[_activeRow].Cells[1].Value = _name;
            dataGridView1.Rows[_activeRow].Cells[2].Value = _description;
            dataGridView1.Rows[_activeRow].Cells[3].Value = _link;

            int[] res = _checkbox_array;
            Array.Sort(res);
            //
            // The CATEGORIES cell should take ints of categories comma separated
            //
            dataGridView1.Rows[_activeRow].Cells[4].Value = string.Join(",", res);

            //
            // Absolutely MUST command before updating dda. No programmatical updates would take place in dataGridView without this method!
            //
            bs.EndEdit();

            try
            {
                    dda = new OleDbDataAdapter("SELECT * FROM [data]", connectionString);
                    cb = new OleDbCommandBuilder(dda);
                    cb.GetUpdateCommand();
                    //
                    // A little trick to get UpdateCommand. No updates without these methods could be possible!
                    //
                    dda.UpdateCommand = cb.GetUpdateCommand();
                    dda.Update(dtab);
                    dtab = new DataTable();
                    dda.Fill(dtab);
                    
            }
            catch (OleDbException exc)
            {
                  MessageBox.Show(exc.Message, "OledbException Error");
            }
            catch (Exception e)
            {
                  MessageBox.Show(e.Message);
            }
          //
          // We don't forget to show the result of update
          //
          ShowData();
        }

        /// <summary>
        /// Button НОВАЯ ПОЗИЦИЯ clicked
        /// </summary>
        /// <param name="sender">Button1 control</param>
        /// <param name="e">Event arguments</param>
        private void button1_Click(object sender, EventArgs e)
        {
            using (OleDbConnection dbConnect = new OleDbConnection(connectionString))
            {
                OleDbCommand command = dbConnect.CreateCommand();
                //
                // To create a new empty row we insert empty values except of id which is AUTOINCREMENTed
                //
                command.CommandText = "insert into data (name, description, link, categories) values ('','','','')";
                try
                {
                    dbConnect.Open();
                    command.ExecuteNonQuery();
                    dataGridView1.DataSource = connectionString;
                    dataGridView1.Refresh();
                    dbConnect.Close();
                }

                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message);

                }

            }
            ShowData();
        }
             
  }
}
