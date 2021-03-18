using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using RevitDatabase;

namespace RevitDatabase_Desktop
{
    public partial class MainForm : System.Windows.Forms.Form
    {
        #region Class Level Variables

        //System.Data
        private DataSet ds = new DataSet();
        private DataTable dt = new DataTable();
        private SqlDataAdapter da;
        public static SqlConnection con = new SqlConnection();

        //System
        private List<string> markValues = new List<string>();
        public static string SelectedDatabase = "";
        public static string SelectedTable = "";
        private string CopiedValue;
        protected StatusBar mainStatusBar = new StatusBar();
        protected StatusBarPanel databasePanel = new StatusBarPanel();
        protected StatusBarPanel tablePanel = new StatusBarPanel();
        private bool tableIsBeingFiltered = false;

        #endregion

        public MainForm()
        {
            InitializeComponent();

            con.ConnectionString = "";

            CreateMainMenu();
            CreateStatusBar();

            tableIsBeingFiltered = false;
            SetMenuItemChecked("filter", false);
            DisableEnableMenuItems();
        }

        #region Voids

        private void GetAllDatabases(System.Windows.Forms.ComboBox _comboBox)
        {
            con.ConnectionString = @"Server=BPRMEPCM-7\SQLEXPRESS;Trusted_Connection=True;";
            con.Open();
            using (SqlCommand com = new SqlCommand("SELECT name from sys.databases", con))
            {
                using (SqlDataReader reader = com.ExecuteReader())
                {
                    _comboBox.Items.Clear();
                    while (reader.Read())
                    {
                        string dbName = "";
                        dbName = (string)reader[0];

                        if (dbName.StartsWith("db"))
                        {
                            _comboBox.Items.Add(dbName);
                        }
                    }
                }
            }
            con.Close();
        }

        private void GetAllTablesFromDatabase(System.Windows.Forms.ComboBox _comboBox)
        {
            con.Open();
            using (SqlCommand com = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", con))
            {
                using (SqlDataReader reader = com.ExecuteReader())
                {
                    _comboBox.Items.Clear();
                    while (reader.Read())
                    {
                        string tblName = "";
                        tblName = (string)reader["TABLE_NAME"];

                        if (tblName.StartsWith("tbl"))
                        {
                            _comboBox.Items.Add(tblName);
                        }
                    }
                }
            }
            con.Close();
        }

        private void SetMenuItemEnabled(string _name, bool _enabled)
        {
            foreach (MenuItem mainMenuItem in Menu.MenuItems)
            {
                foreach (MenuItem subMenuItem in mainMenuItem.MenuItems)
                {
                    if (subMenuItem.Name == _name)
                        subMenuItem.Enabled = _enabled;
                }
            }
        }

        private void SetMenuItemChecked(string _name, bool _checked)
        {
            foreach (MenuItem mainMenuItem in Menu.MenuItems)
            {
                foreach (MenuItem subMenuItem in mainMenuItem.MenuItems)
                {
                    if (subMenuItem.Name == _name)
                        subMenuItem.Checked = _checked;
                }
            }
        }

        private void UpdateTable(DataGridView _dgv, string _tableName)
        {
            if (MessageBox.Show("Are you sure you want to update the database table?", "Update Table", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                foreach (DataGridViewRow row in _dgv.Rows)
                {
                    int currentDictIndex = 0;

                    SqlCommand command = new SqlCommand();
                    command.Connection = con;
                    command.CommandType = CommandType.Text;
                    command.Connection = con;
                    command.CommandType = CommandType.Text;

                    StringBuilder commandText = new StringBuilder();
                    commandText.Append("UPDATE [" + _tableName + "] SET ");

                    con.Open();
                    Dictionary<string, string> dict = new Dictionary<string, string>();
                    dict = CellValueDictionary(row);

                    foreach (string param in dict.Keys)
                    {
                        string mark = "";
                        mark = row.Cells["Mark"].Value.ToString();
                        string parameterValue = "";
                        parameterValue = dict[param];
                        command.Parameters.AddWithValue("@" + param, parameterValue + "");

                        if (currentDictIndex < dict.Count - 1)
                        {
                            commandText.Append(param + "=@" + param + ",");
                        }
                        else if (currentDictIndex == dict.Count - 1)
                        {
                            commandText.Append(param + "=@" + param + " WHERE Mark=\'" + mark + "\'");
                        }

                        currentDictIndex++;
                    }

                    command.CommandText = commandText.ToString();
                    command.ExecuteNonQuery();
                    con.Close();
                }

                if (!tableIsBeingFiltered)
                    LoadTable(_tableName);
                else
                    FilterTable(_tableName);
            }
        }

        private void FillDataGridView()
        {

            DrawingControl.SetDoubleBuffered(dgvData);
            DrawingControl.SuspendDrawing(dgvData);

            dgvData.DataSource = null;

            dgvData.Columns.Clear();

            dt.Columns.Clear();
            dt.Clear();

            da.Fill(dt);

            dgvData.DataSource = dt.DefaultView;

            dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
            dgvData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvData.ReadOnly = false;

            //ID
            dgvData.Columns["ID"].Visible = false;
            dgvData.Columns["ID"].ReadOnly = true;

            //ElementId
            dgvData.Columns["ElementId"].Visible = false;
            dgvData.Columns["ElementId"].ReadOnly = true;

            //Mark
            dgvData.Columns["Mark"].Visible = true;
            dgvData.Columns["Mark"].ReadOnly = true;

            dgvData.ClearSelection();

            DrawingControl.ResumeDrawing(dgvData);

        }

        private void LoadTable(string _tableName)
        {
            tableIsBeingFiltered = false;
            con.Open();
            string sql = "SELECT * FROM [" + _tableName + "]";

            FillDataSet(sql, _tableName);
            FillDataGridView();

            con.Close();
            markValues.Clear();

            foreach (DataGridViewRow row in dgvData.Rows)
            {
                markValues.Add(row.Cells["Mark"].Value.ToString());
            }

            dgvData.ClearSelection();
        }

        private void SQLCommand(string _sql)
        {
            SqlCommand command = new SqlCommand(_sql);
            command.Connection = con;
            command.ExecuteNonQuery();
        }

        private void FilterTable(string _tableName)
        {
            tableIsBeingFiltered = true;
            List<string> tableFields = new List<string>();

            con.Open();
            DataColumnCollection col;
            col = dt.Columns;

            foreach (DataColumn column in col)
            {
                if (column.ColumnName != "ID")
                {
                    string param = column.ColumnName;
                    tableFields.Add(param);
                }
            }

            string searchString = txtFilter.Text;

            DrawingControl.SetDoubleBuffered(dgvData);
            DrawingControl.SuspendDrawing(dgvData);

            dgvData.DataSource = null;
            dgvData.Columns.Clear();

            DataView dsView = new DataView();
            dsView = ds.Tables[0].DefaultView;

            BindingSource bs = new BindingSource();
            bs.DataSource = dsView;

            string filterString = FilterLikeString(tableFields, searchString);
            bs.Filter = filterString;

            dgvData.DataSource = bs;

            con.Close();

            dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
            dgvData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvData.ReadOnly = false;

            //ID
            dgvData.Columns["ID"].Visible = false;
            dgvData.Columns["ID"].ReadOnly = true;

            //ElementId
            dgvData.Columns["ElementId"].Visible = false;
            dgvData.Columns["ElementId"].ReadOnly = true;

            //Mark
            dgvData.Columns["Mark"].Visible = true;
            dgvData.Columns["Mark"].ReadOnly = true;

            dgvData.ClearSelection();

            DrawingControl.ResumeDrawing(dgvData);
        }

        private void ClearValues()
        {
            if (MessageBox.Show("Are you sure you want to clear the values of the selected cells?", "Clear Values", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                foreach (DataGridViewCell cell in dgvData.SelectedCells)
                {
                    cell.Value = "";
                }
            }
        }

        private void PasteValue()
        {
            if (MessageBox.Show("Are you sure you want to paste the value below to the selected cells?", "Paste Values", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                foreach (DataGridViewCell cell in dgvData.SelectedCells)
                {
                    cell.Value = CopiedValue;
                }
            }
        }

        private void CreateStatusBar()
        {
            databasePanel.BorderStyle = StatusBarPanelBorderStyle.Sunken;
            databasePanel.Text = "Database:";
            databasePanel.AutoSize = StatusBarPanelAutoSize.Spring;
            mainStatusBar.Panels.Add(databasePanel);

            tablePanel.BorderStyle = StatusBarPanelBorderStyle.Sunken;
            tablePanel.Text = "Table:";
            tablePanel.AutoSize = StatusBarPanelAutoSize.Spring;
            mainStatusBar.Panels.Add(tablePanel);

            mainStatusBar.ShowPanels = true;

            this.Controls.Add(mainStatusBar);
        }

        private void CreateMainMenu()
        {
            MainMenu mainMenu = new MainMenu();

            #region FileMenu

            MenuItem fileMenu = new MenuItem();
            MenuItem menuSelectDatabase = new MenuItem();
            MenuItem menuCreateTable = new MenuItem();
            MenuItem menuLoadTable = new MenuItem();
            MenuItem menuCloseTable = new MenuItem();
            MenuItem menuExit = new MenuItem();

            menuSelectDatabase.Name = "select_database";
            menuCreateTable.Name = "create_table";
            menuLoadTable.Name = "load_table";
            menuCloseTable.Name = "close_table";

            fileMenu.Text = "File";

            menuSelectDatabase.Text = "Select Database\u2026";
            menuCreateTable.Text = "Create Table\u2026";
            menuLoadTable.Text = "Load Table\u2026";
            menuCloseTable.Text = "Close Table";

            menuExit.Text = "Exit";

            fileMenu.MenuItems.Add(menuSelectDatabase);
            fileMenu.MenuItems.Add("-");
            fileMenu.MenuItems.Add(menuCreateTable);
            fileMenu.MenuItems.Add(menuLoadTable);
            fileMenu.MenuItems.Add(menuCloseTable);
            fileMenu.MenuItems.Add("-");
            fileMenu.MenuItems.Add(menuExit);

            menuSelectDatabase.Click += new System.EventHandler(this.menuSelectDatabase_Click);
            menuLoadTable.Click += new System.EventHandler(this.menuLoadTable_Click);
            menuCloseTable.Click += new System.EventHandler(this.menuCloseTable_Click);
            menuExit.Click += new System.EventHandler(this.menuExit_Click);

            #endregion

            #region EditMenu

            MenuItem editMenu = new MenuItem();
            MenuItem menuUpdateTable = new MenuItem();
            MenuItem menuFilter = new MenuItem();

            menuUpdateTable.Name = "update_table";
            menuFilter.Name = "filter";

            editMenu.Text = "Edit";

            menuUpdateTable.Text = "Update Table";
            menuFilter.Text = "Filter";

            editMenu.MenuItems.Add(menuUpdateTable);
            editMenu.MenuItems.Add("-");
            editMenu.MenuItems.Add(menuFilter);

            menuUpdateTable.Click += new System.EventHandler(this.menuUpdateTable_Click);
            menuFilter.Click += new System.EventHandler(this.menuFilter_Click);

            #endregion

            mainMenu.MenuItems.Add(fileMenu);
            mainMenu.MenuItems.Add(editMenu);

            this.Menu = mainMenu;

        }

        private void DisableEnableMenuItems()
        {
            if (SelectedDatabase == "" && SelectedTable == "")
            {
                //file menu
                SetMenuItemEnabled("select_table", true);
                SetMenuItemEnabled("create_table", false);
                SetMenuItemEnabled("load_table", false);
                SetMenuItemEnabled("close_table", false);

                //edit menu
                SetMenuItemEnabled("update_table", false);
                SetMenuItemEnabled("filter", false);

                txtFilter.Text = "";
                txtFilter.Enabled = false;
                btnFilter.Enabled = false;
                btnClear.Enabled = false;

            }
            else if (SelectedDatabase != "" && SelectedTable == "")
            {
                //file menu
                SetMenuItemEnabled("select_table", true);
                SetMenuItemEnabled("create_table", true);
                SetMenuItemEnabled("load_table", true);
                SetMenuItemEnabled("close_table", false);

                //edit menu
                SetMenuItemEnabled("update_table", false);
                SetMenuItemEnabled("filter", false);

                txtFilter.Text = "";
                txtFilter.Enabled = false;
                btnFilter.Enabled = false;
                btnClear.Enabled = false;

            }
            else if (SelectedDatabase != "" && SelectedTable != "" && tableIsBeingFiltered)
            {
                //file menu
                SetMenuItemEnabled("select_table", false);
                SetMenuItemEnabled("create_table", false);
                SetMenuItemEnabled("load_table", false);
                SetMenuItemEnabled("close_table", false);

                //edit menu
                SetMenuItemEnabled("update_table", true);

                txtFilter.Text = "";
                txtFilter.Enabled = true;
                btnFilter.Enabled = true;
                btnClear.Enabled = true;
            }
            else
            {
                //file menu
                SetMenuItemEnabled("select_table", true);
                SetMenuItemEnabled("create_table", true);
                SetMenuItemEnabled("load_table", true);
                SetMenuItemEnabled("close_table", true);

                //edit menu
                SetMenuItemEnabled("update_table", true);
                SetMenuItemEnabled("filter", true);

                txtFilter.Text = "";
                txtFilter.Enabled = false;
                btnFilter.Enabled = false;
                btnClear.Enabled = false;
            }
        }

        #endregion

        #region Functions

        private Dictionary<string, string> CellValueDictionary(DataGridViewRow _row)
        {
            DataColumnCollection col;
            col = dt.Columns;
            Dictionary<string, string> dict = new Dictionary<string, string>();

            foreach (DataColumn column in col)
            {
                if (column.ColumnName != "ID" && column.ColumnName != "ElementId" && column.ColumnName != "Mark")
                {
                    string columnName = "";
                    columnName = column.ColumnName;
                    string cellValue = "";
                    cellValue = _row.Cells[columnName].Value.ToString();
                    dict.Add(columnName, cellValue);
                }

            }

            return dict;
        }

        private DataSet FillDataSet(string _sql, string _tableName)
        {
            da = new SqlDataAdapter(_sql, con);
            ds.Tables.Clear();
            ds.Clear();
            da.Fill(ds, "[" + _tableName + "]");

            return ds;
        }

        private string FilterLikeString(List<string> _tableFields, string _searchString)
        {
            string filterString = "";
            StringBuilder likeFilter = new StringBuilder();
            int counter = 0;

            foreach (string field in _tableFields)
            {
                int fieldCount = _tableFields.Count - 1;

                if (counter < fieldCount)
                    likeFilter.Append(field + " LIKE'%" + _searchString + "%' or ");
                else
                    likeFilter.Append(field + " LIKE'%" + _searchString + "%'");

                counter += 1;
            }

            filterString = likeFilter.ToString();

            return filterString;
        }

        private ContextMenu TableContextMenu()
        {
            ContextMenu mnu = new ContextMenu();
            MenuItem cxmnuCopyValue = new MenuItem("Copy Value");
            MenuItem cxmnuPasteValue = new MenuItem("Paste Value");
            MenuItem cxmnuClearValues = new MenuItem("Clear Values");

            cxmnuClearValues.Click += new EventHandler(cxmnuClearValues_Click);
            cxmnuCopyValue.Click += new EventHandler(cxmnuCopyValue_Click);
            cxmnuPasteValue.Click += new EventHandler(cxmnuPasteValue_Click);

            mnu.MenuItems.Add(cxmnuCopyValue);
            mnu.MenuItems.Add(cxmnuPasteValue);
            mnu.MenuItems.Add("-");
            mnu.MenuItems.Add(cxmnuClearValues);

            return mnu;
        }

        #endregion

        #region MainMenuEvents

        private void menuUpdateTable_Click(object sender, EventArgs e)
        {
            UpdateTable(dgvData, SelectedTable);
        }

        private void menuSelectDatabase_Click(object sender, EventArgs e)
        {
            frmSelectionBox new_frmSelectionBox = new frmSelectionBox();
            new_frmSelectionBox.Text = "Select Database";
            new_frmSelectionBox.lblInstructions.Text = "Select the database you want to connect to\nfrom the drop-down list below";

            System.Windows.Forms.ComboBox cbDatabases;
            cbDatabases = new_frmSelectionBox.cbItems;

            GetAllDatabases(cbDatabases);

            if (new_frmSelectionBox.ShowDialog() == DialogResult.OK)
            {
                con.ConnectionString = "";
                SelectedDatabase = "";
                SelectedDatabase = cbDatabases.SelectedItem.ToString();
                con.ConnectionString = @"Server=BPRMEPCM-7\SQLEXPRESS;Database=" + SelectedDatabase + @";Trusted_Connection=True;";
                databasePanel.Text = "Database: " + SelectedDatabase;
                tablePanel.Text = "Table:";
                tableIsBeingFiltered = false;
                DisableEnableMenuItems();
            }
        }

        private void menuLoadTable_Click(object sender, EventArgs e)
        {
            frmSelectionBox new_frmSelectionBox = new frmSelectionBox();
            new_frmSelectionBox.Text = "Load Table";
            new_frmSelectionBox.lblInstructions.Text = "Select the table you want to load\nfrom the drop-down list below";

            System.Windows.Forms.ComboBox cbTables;
            cbTables = new_frmSelectionBox.cbItems;

            if (SelectedDatabase != "")
                GetAllTablesFromDatabase(cbTables);

            if (new_frmSelectionBox.ShowDialog() == DialogResult.OK)
            {
                SelectedTable = cbTables.SelectedItem.ToString();
                LoadTable(SelectedTable);
                databasePanel.Text = "Database: " + SelectedDatabase;
                tablePanel.Text = "Table: " + SelectedTable;
                DisableEnableMenuItems();
            }
        }

        private void menuCloseTable_Click(object sender, EventArgs e)
        {
            dgvData.DataSource = null;
            dgvData.Columns.Clear();
            dt.Columns.Clear();
            dt.Clear();
            SelectedTable = "";
            tableIsBeingFiltered = false;
            databasePanel.Text = "Database: " + SelectedDatabase;
            tablePanel.Text = "Table:";
            DisableEnableMenuItems();
        }

        private void menuExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #endregion

        #region Button Events

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtFilter.Text = "";
            FilterTable(SelectedTable);
            tableIsBeingFiltered = false;
        }

        private void btnFilter_Click(object sender, EventArgs e)
        {
            FilterTable(SelectedTable);
        }

        #region ContextMenu

        private void cxmnuClearValues_Click(object sender, EventArgs e)
        {
            ClearValues();
        }

        private void cxmnuCopyValue_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in dgvData.SelectedCells)
            {
                if (dgvData.SelectedCells.Count == 1)
                    CopiedValue = cell.Value.ToString();
                else
                    return;
            }
        }

        private void cxmnuPasteValue_Click(object sender, EventArgs e)
        {
            PasteValue();
        }

        #endregion

        #endregion

        private void dgvData_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenu contextMenu = new ContextMenu();
                contextMenu = TableContextMenu();
                contextMenu.Show(dgvData, new System.Drawing.Point(e.X, e.Y));
            }

        }

        private void menuFilter_Click(object sender, EventArgs e)
        {
            tableIsBeingFiltered = !tableIsBeingFiltered;
            SetMenuItemChecked("filter", tableIsBeingFiltered);

            DisableEnableMenuItems();
        }

    }

    #region DrawingControl

    public static class DrawingControl
    {
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr _hWnd, Int32 _wMsg, bool _wParam, Int32 _lParam);

        private const int WM_SETREDRAW = 11;

        public static void SetDoubleBuffered(System.Windows.Forms.Control _ctrl)
        {
            if (!SystemInformation.TerminalServerSession)
            {
                typeof(System.Windows.Forms.Control).InvokeMember("DoubleBuffered", (System.Reflection.BindingFlags.SetProperty
                                | (System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)), null, _ctrl, new object[] {
                            true});
            }
        }

        public static void SetDoubleBuffered_ListControls(List<System.Windows.Forms.Control> _ctrlList)
        {
            if (!SystemInformation.TerminalServerSession)
            {
                foreach (System.Windows.Forms.Control ctrl in _ctrlList)
                {
                    typeof(System.Windows.Forms.Control).InvokeMember("DoubleBuffered", (System.Reflection.BindingFlags.SetProperty
                                    | (System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)), null, ctrl, new object[] {
                                true});
                }
            }
        }

        public static void SuspendDrawing(System.Windows.Forms.Control _ctrl)
        {
            SendMessage(_ctrl.Handle, WM_SETREDRAW, false, 0);
        }

        public static void SuspendDrawing_ListControls(List<System.Windows.Forms.Control> _ctrlList)
        {
            foreach (System.Windows.Forms.Control ctrl in _ctrlList)
            {
                SendMessage(ctrl.Handle, WM_SETREDRAW, false, 0);
            }
        }

        public static void ResumeDrawing(System.Windows.Forms.Control _ctrl)
        {
            SendMessage(_ctrl.Handle, WM_SETREDRAW, true, 0);
            _ctrl.Refresh();
        }

        public static void ResumeDrawing_ListControls(List<System.Windows.Forms.Control> _ctrlList)
        {
            foreach (System.Windows.Forms.Control ctrl in _ctrlList)
            {
                SendMessage(ctrl.Handle, WM_SETREDRAW, true, 0);
                ctrl.Refresh();
            }
        }
    }

    #endregion

}
