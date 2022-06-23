using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Xml.Serialization;

namespace Booktable
{
    public partial class MainForm : Form
    {
        string raw_file_path = "";
        DirectoryInfo source_dir = new DirectoryInfo(Path.Combine(Environment.CurrentDirectory, @"source"));
        string data_file_path = null;

        List<List<string>> dataList = new List<List<string>>();
        List<string> columnName = new List<string>();

        List<List<string>> viewList = new List<List<string>>();
        List<string> viewColumnName = new List<string>();

        // history for save.
        List<int> history = new List<int>();
        List<List<string>> buyer_info = new List<List<string>>();

        // history for cancle the last change (maximum save = 10 changes).
        int bufferCapacity = 10;
        List<int> bufferhistory = new List<int>();

        // search word.
        string search_word = "";

        // Main workbook for new excel file loading.
        private Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();

        public MainForm()
        {
            InitializeComponent();

            this.FormClosing += Main_FormClosing;
            this.button1.Click += button1_Click;
            this.button2.Click += button2_Click;
            this.button3.Click += button3_Click;
            this.button5.Click += button5_Click;
            this.button14.Click += button14_Click;
            this.button16.Click += button16_Click;
            this.tabControl1.MouseClick += TabControl1_MouseClick;
            this.textBox1.TextChanged += textBox1_TextChanged;
            this.checkBox1.CheckedChanged += CheckBox1_CheckedChanged;
            this.checkBox2.CheckedChanged += CheckBox2_CheckedChanged;

            this.panel4.DragEnter += new DragEventHandler(Panel4_DragEnter);
            this.panel4.DragDrop += new DragEventHandler(Panel4_DragDrop);

            this.dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
            this.dataGridView2.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView2.RowHeadersVisible = false;


            // Load recent data excel file if it exist.
            this.GetMostRecentDataExcel();
        }

        private void DataListToViewList()
        {
            int IsSoldIdx = this.columnName.IndexOf("판매여부");

            this.viewList = new List<List<string>>(); // initialize
            this.viewColumnName = this.columnName.ToList();
            this.viewColumnName.Add("남은권수");
            this.viewColumnName.Add("팔린권수");
            this.viewColumnName.Add("총권수");

            // each row
            foreach (List<string> oneRow in this.dataList)
            {
                string title = oneRow[0];

                // If the title is not added to viewlist,
                if (this.viewList.Where(item => item[0]==title).Count() == 0)
                {
                    List<string> tmp = oneRow.ToList(); //deep-copy
                    int NotSoldNum = this.dataList.Where(item => item[0] == title && item[IsSoldIdx] == "X").Count();
                    int SoldNum = this.dataList.Where(item => item[0] == title).Count() - NotSoldNum;

                    tmp.Add(NotSoldNum.ToString()); // number of books not sold
                    tmp.Add(SoldNum.ToString()); // number of books sold
                    tmp.Add((NotSoldNum + SoldNum).ToString()); // number of total books
                    this.viewList.Add(tmp);
                }
            }
        }

        private void TabControl1_MouseClick(object sender, MouseEventArgs e)
        {
            if (this.tabControl1.SelectedIndex==1)
            {
                this.UpdateDailyChart();
                this.UpdateTotalLabel();
                this.UpdateRatios();
            }

            if (this.tabControl1.SelectedIndex==2)
            {
                // Load file tree
                FormTree_Load();
            }
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult response = MessageBox.Show("정말로 종료하시겠습니까? 데이터는 자동으로 저장됩니다.", "프로그램 종료", MessageBoxButtons.YesNo);
            if ( response == DialogResult.No)
            {
                e.Cancel = true;
            }

            else
            {
                this.SaveMemo();
                this.Save_CurrentExcel();
                this.application.Quit();
            }
        }

        private void GetMostRecentDataExcel()
        {
            // If the most recent data file has the correct file extension, Load it.

            try
            {
                string mostrecentfile = this.source_dir.GetFiles()
                                        .OrderByDescending(f => f.LastWriteTime)
                                        .Where(f => Path.GetExtension(f.Name).ToUpper() == ".XLSX")
                                        .First()
                                        .ToString();

                this.data_file_path = Path.Combine(Environment.CurrentDirectory, @"source\" + mostrecentfile);

                if (this.data_file_path.EndsWith(".xlsx"))
                {
                    this.LoadExcelFile(this.data_file_path);
                    this.Update_dataGridViews();
                }
            }

            // IF file does not exist, ignore the error and continue.
            catch { this.data_file_path = null; }
        }

        private void LoadExcelFile(string file_path = null)
        {
            try
            {
                Workbook workbook;

                // initialize the dataList and colnames.
                dataList = new List<List<string>>();
                columnName = new List<string>(); // column name of datalist

                if (file_path != null)
                {
                    workbook = this.application.Workbooks.Open(Filename: file_path);
                }
                else
                {
                    workbook = this.application.Workbooks.Open(Filename: this.raw_file_path);
                }


                _Worksheet xlWorksheet = workbook.Sheets[1];
                Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    List<string> tmp = new List<string>(); // row 

                    for (int j = 1; j <= colCount; j++)
                    {
                        string item = "-";

                        try
                        {
                            item = xlRange.Cells[i, j].Value2.ToString();
                        }

                        catch
                        {
                            item = xlRange.Cells[i, j].Value2;
                        }

                        // header (column name) row
                        if (i == 1)
                        {
                            this.columnName.Add(item);
                        }
                        // other rows
                        else
                        {
                            tmp.Add(item);
                        }
                    }

                    if (i != 1)
                    {
                        this.dataList.Add(tmp.ToList());
                    }
                }
                this.textBox2.Clear();
                workbook.Close();

                this.DataListToViewList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " !");
                this.textBox2.Clear();
                this.application.Quit();
            }
        }

        private void OpenExcelFile(string file_path = null)
        {
            try
            {
                Workbook workbook;

                // initialize the dataList and colnames.
                dataList = new List<List<string>>();
                columnName = new List<string>(); // column name of datalist

                if (file_path != null)
                {
                    workbook = this.application.Workbooks.Open(Filename: file_path);
                }
                else
                {
                    workbook = this.application.Workbooks.Open(Filename: this.raw_file_path);
                }


                _Worksheet xlWorksheet = workbook.Sheets[1];
                Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    List<string> tmp = new List<string>(); // row 

                    for (int j = 1; j <= colCount; j++)
                    {
                        string item = "-";

                        try
                        {
                            item = xlRange.Cells[i, j].Value2.ToString();
                        }

                        catch
                        {
                            item = xlRange.Cells[i, j].Value2;
                        }

                        // ignore '권수' as column name. It will be reflected by number of rows. 
                        if (j != colCount)
                        {
                            // header (column name) row
                            if (i == 1)
                            {
                                this.columnName.Add(item);
                            }
                            // other rows
                            else
                            {
                                tmp.Add(item);
                            }
                        }
                    }

                    if (i != 1)
                    {
                        //
                        int num_books = Int32.Parse(xlRange.Cells[i, colCount].Value2.ToString());
                        for (int k = 0; k < num_books; k++)
                        {
                            this.dataList.Add(tmp.ToList());
                        }
                    }
                }

                // Check whether the excel file contains IsSold, WhenSold, HowSold.
                string check1 = "판매여부";
                string check2 = "판매시간";
                string check3 = "결제방법";
                string check4 = "구매자";

                if (this.columnName.Contains(check1) && this.columnName.Contains(check2) && this.columnName.Contains(check3) && this.columnName.Contains(check4))
                {
                    // do nothing
                }

                else
                {
                    // This may be excuted only for the first excel file open.
                    // Add two thing together
                    this.columnName.Add(check1);
                    this.columnName.Add(check2);
                    this.columnName.Add(check3);
                    this.columnName.Add(check4);

                    foreach (List<string> oneRow in this.dataList)
                    {
                        // Add initial value for IsSold.
                        oneRow.Add("X");
                        // Add initial value for WhenSold.
                        oneRow.Add("-");
                        // Add initial value for HowSold.
                        oneRow.Add("-");
                        // Add initial value for WhoBuy.
                        oneRow.Add("-");
                    }

                    // Now, save it as new (actually, copy version) excel file so that the program changes it only after this.
                    this.Save_CurrentExcel();
                }
                this.textBox2.Clear();
                workbook.Close();

                this.DataListToViewList();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + " !");
                this.application.Quit();
            }

        }


        private void Save_CurrentExcel()  // Do not change all thing. Keep changed history and change it only.
        {
            // If data Excel does not exist, create new and save it.
            if (this.data_file_path == null)
            {
                // First, copy raw excel file into source directory.
                DateTime currentTime = DateTime.Now;
                string save_file_name = "DataExcel_" + currentTime.ToString("yyyyMMddHHmmss") + ".xlsx";
                string source_dir = Path.Combine(Environment.CurrentDirectory, @"source");

                try
                {
                    File.Copy(this.raw_file_path, source_dir + '/' + save_file_name);

                    // Second, change copied excel file.
                    // Open
                    Workbook workbook = this.application.Workbooks.Open(source_dir + '/' + save_file_name);
                    _Worksheet xlWorksheet = workbook.Sheets[1];
                    Range xlRange = xlWorksheet.UsedRange;
                    xlWorksheet.Columns.NumberFormat = "@";

                    // Change
                    for (int j=0; j<this.columnName.Count; j++)
                    {
                        xlWorksheet.Cells[1, j + 1] = this.columnName[j];
                        for (int i=0; i<this.dataList.Count; i++)
                        {
                            xlWorksheet.Cells[i + 2, j + 1] = this.dataList[i][j];
                        }
                    }

                    // Save
                    workbook.Save();
                    workbook.Close(true, Type.Missing, Type.Missing);

                }
                catch (Exception ex)
                {
                    //workbook.Close(true, Type.Missing, Type.Missing);
                    MessageBox.Show(ex.Message + " !");
                    this.application.Quit();
                }
               
                this.data_file_path = source_dir + '/' + save_file_name;
            }

            // If data Excel is already opened, change it.
            else
            {
                try
                {
                    Workbook workbook = this.application.Workbooks.Open(this.data_file_path);
                    _Worksheet xlWorksheet = workbook.Sheets[1];
                    Range xlRange = xlWorksheet.UsedRange;
                    xlWorksheet.Columns.NumberFormat = "@";

                    // Change
                    for (int j = 0; j < this.dataList[0].Count; j++)
                    {
                        // fill column head.
                        xlWorksheet.Cells[1, j + 1] = this.columnName[j];
                    }

                    foreach (int i in this.history)
                    {
                        for (int j = 1; j <= this.dataList[0].Count; j++)
                        {
                            xlWorksheet.Cells[i + 2, j] = this.dataList[i][j - 1];
                        }
                    }

                    //save
                    workbook.Save();
                    workbook.Close(true, Type.Missing, Type.Missing);

                    this.history = new List<int>();
                    this.buyer_info = new List<List<string>>();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + " !");
                    this.application.Quit();
                }
            }

            // if all things are done, show message on the statusstrip
            DateTime saveTime = DateTime.Now;
            this.toolStripStatusLabel1.Text = "현재 데이터를 저장했습니다 (" + saveTime.ToString("yyyy/MM/dd HH:mm:ss") + ").";
        }

        private void SellBook(string who, string optional, List<string> targetBook = null)
        {
            try
            {
                int IsSoldIdx = this.columnName.IndexOf("판매여부");
                int WhenSoldIdx = this.columnName.IndexOf("판매시간");
                int HowSoldIdx = this.columnName.IndexOf("결제방법");
                int WhoBuyIdx = this.columnName.IndexOf("구매자");
                int OptionalIdx = this.columnName.IndexOf("비고");


                List<string> lastSelectedBookInList = new List<string>();

                if (targetBook != null)
                {
                    lastSelectedBookInList = targetBook;
                }
                else
                {
                    // Find the last book whose name is 'selectedBook' in the dataList.
                    string selectedBook = this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                    lastSelectedBookInList = this.dataList.FindLast(
                        delegate (List<string> oneElement)
                        {
                            return (oneElement[0] == selectedBook) && (oneElement[IsSoldIdx]=="X");
                        });
                }


                // Change the selected book into sold.
                DateTime currentTime = DateTime.Now;
                lastSelectedBookInList[IsSoldIdx] = "O";
                lastSelectedBookInList[WhenSoldIdx] = currentTime.ToString("yyyy/MM/dd HH:mm:ss");
                lastSelectedBookInList[WhoBuyIdx] = who;
                lastSelectedBookInList[OptionalIdx] = optional;

                int checkednum = this.CheckCheckbox();

                switch (checkednum)
                {
                    case 1:
                        lastSelectedBookInList[HowSoldIdx] = "현금";
                        break;
                    case 2:
                        lastSelectedBookInList[HowSoldIdx] = "계좌이체";
                        break;
                }

                // Update viewList
                this.DataListToViewList();

                // Update datagridview
                this.Update_dataGridViews();

                // Update history
                if (targetBook == null)
                {
                    int changedIdx = this.dataList.IndexOf(lastSelectedBookInList);
                    this.history.Add(changedIdx);
                    this.Update_bufferhistory(changedIdx);
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " !");
            }
        }

        private int CheckCheckbox()
        {
            if (this.checkBox1.Checked && !this.checkBox2.Checked)
            {
                return 1;
            }
            else if (!this.checkBox1.Checked && this.checkBox2.Checked)
            {
                return 2;
            }
            else { return 0; }
        }

        private void CancleBook(List<string> targetBook = null)
        {
            try
            {
                int IsSoldIdx = this.columnName.IndexOf("판매여부");
                int WhenSoldIdx = this.columnName.IndexOf("판매시간");
                int WhenSoldCell = this.dataGridView2.Columns["판매시간"].Index;
                int HowSoldIdx = this.columnName.IndexOf("결제방법");
                int WhoBuyIdx = this.columnName.IndexOf("구매자");
                int OptionalIdx = this.columnName.IndexOf("비고");

                string selectedBook = this.dataGridView2.SelectedRows[0].Cells[0].Value.ToString();
                string selectedSoldTime = this.dataGridView2.SelectedRows[0].Cells[WhenSoldCell].Value.ToString();

                List<string> lastSelectedBookInList = new List<string>();

                if (targetBook != null)
                {
                    lastSelectedBookInList = targetBook;
                }
                else
                {
                    // Find the last book whose name is 'selectedBook' in the dataList.
                    lastSelectedBookInList = this.dataList.FindLast(
                        delegate (List<string> oneElement)
                        {
                            return oneElement[0] == selectedBook && oneElement[WhenSoldIdx].Equals(selectedSoldTime);
                        });
                }

                // Add buyer information to history
                List<string> buyer_info = new List<string> { lastSelectedBookInList[WhoBuyIdx], 
                                                                lastSelectedBookInList[OptionalIdx] };
                this.buyer_info.Add(buyer_info);

                // Change the selected book into unsold.
                DateTime currentTime = DateTime.Now;
                lastSelectedBookInList[IsSoldIdx] = "X";
                lastSelectedBookInList[WhenSoldIdx] = "-";
                lastSelectedBookInList[HowSoldIdx] = "-";
                lastSelectedBookInList[WhoBuyIdx] = "-";
                lastSelectedBookInList[OptionalIdx] = "-";

                // update viewList
                this.DataListToViewList();

                // Update datagridview
                this.Update_dataGridViews();

                if (targetBook == null)
                {
                    int changedIdx = this.dataList.IndexOf(lastSelectedBookInList);
                    if (this.history.Contains(changedIdx))
                    {
                        this.history.Remove(changedIdx);
                    }
                    else
                    {
                        this.history.Add(changedIdx);
                    }
                    this.Update_bufferhistory(changedIdx);
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " !");
            }
        }

        private void Update_bufferhistory(int newidx)
        {
            if (this.bufferhistory.Count < this.bufferCapacity)
            {
                this.bufferhistory.Add(newidx);
            }

            else
            {
                // Remove the most previous one.
                this.bufferhistory.RemoveAt(0);
                // new one.
                this.bufferhistory.Add(newidx);
            }
        }

        private int CalculateTotal()
        {
            int IsSoldIdx = this.columnName.IndexOf("판매여부");
            int PriceIdx = this.columnName.IndexOf("판매가");

            int total = 0;
            foreach (List<string> oneRow in this.dataList)
            {
                if (oneRow[IsSoldIdx] == "O")
                {
                    total += Convert.ToInt32(oneRow[PriceIdx]);
                }
            }

            return total;
        }

        private void UpdateTotalLabel()
        {
            int total = this.CalculateTotal();
            this.label14.Text = Convert.ToString(total) + " 원";
        }

        private void UpdateRatios()
        {
            int IsSoldIdx = this.columnName.IndexOf("판매여부");
            int HowSoldIdx = this.columnName.IndexOf("결제방법");

            double accountsold = 0;
            double cashsold = 0;
            double total = 0;

            foreach (List<string> oneRow in this.dataList)
            {
                total++;
                if (oneRow[IsSoldIdx] == "O")
                {
                    if (oneRow[HowSoldIdx]=="현금")
                    {
                        cashsold++;
                    }
                    else if (oneRow[HowSoldIdx]=="계좌이체")
                    {
                        accountsold++;
                    }
                }
            }

            double totalsoldratio = (cashsold + accountsold) / (total);
            double accountratio = (accountsold) / (cashsold + accountsold);
            double cashratio = (cashsold) / (cashsold + accountsold);
            this.progressBar1.Value = (int)(totalsoldratio * 100);
            this.label12.Text = String.Format("{0:0.0}%", totalsoldratio*100);
        }

        private void CancleLastUpdate()
        {
            if (this.bufferhistory.Count > 0)
            {
                int WhenSoldIdx = this.columnName.IndexOf("판매시간");

                try
                {
                    List<string> lastChangedBook = this.dataList[this.bufferhistory[this.bufferhistory.Count - 1]];
                    // If last update was 'cancle'
                    if (lastChangedBook[WhenSoldIdx].Equals("-"))
                    {
                        List<string> lastBuyer = this.buyer_info[this.buyer_info.Count - 1];
                        this.buyer_info.RemoveAt(this.buyer_info.Count - 1);
                        this.SellBook(who: lastBuyer[0], optional: lastBuyer[1], targetBook: lastChangedBook);
                    }
                    // If last update was 'sell'
                    else
                    {
                        this.CancleBook(targetBook: lastChangedBook);
                    }

                    this.bufferhistory.RemoveAt(this.bufferhistory.Count - 1);
                }
                catch
                {

                }
                
            }
        }

        private string RemoveAllWhiteSpace(string instr)
        {
            //return String.Concat(instr.Where(c => !Char.IsWhiteSpace(c)));
            try
            {
                return Regex.Replace(instr, @"\s+", "");
            }
            catch
            {
                return "";
            }
            
        }

        private void Update_dataGridView1()
        {   
            int IsSoldIdx = this.columnName.IndexOf("판매여부");
            int UnsoldNumBook = this.viewColumnName.IndexOf("남은권수");
            int SoldNumBook = this.viewColumnName.IndexOf("팔린권수");
            int TotalNumBook = this.viewColumnName.IndexOf("총권수");
            int WhenSoldIdx = this.columnName.IndexOf("판매시간");

            List<string> displayColumnList = new List<string> {"책이름", "저자", "정가", "판매가", "남은권수"};

            // Add column headers.
            foreach (string colname in displayColumnList)
            {
                DataGridViewColumn column = new DataGridViewTextBoxColumn();
                column.Name = colname;
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                column.SortMode = DataGridViewColumnSortMode.Automatic;          
                dataGridView1.Columns.Add(column);
            }

            // Add contents (each row) if the current row has IsSold=="X".
            string search_nospace = this.RemoveAllWhiteSpace(this.search_word);
            int properRowNum = this.viewList.Count(item => Int32.Parse(item[UnsoldNumBook]) > 0 && (this.RemoveAllWhiteSpace(item[0]).Contains(search_nospace) || this.RemoveAllWhiteSpace(item[1]).Contains(search_nospace)));

            if (properRowNum > 0)
            {
                // Second, Add rows.
                dataGridView1.RowCount = properRowNum;

                int rowIdx = 0;
                foreach (List<string> oneRow in this.viewList)
                {
                    if (Int32.Parse(oneRow[UnsoldNumBook]) > 0 && (this.RemoveAllWhiteSpace(oneRow[0]).Contains(search_nospace) || this.RemoveAllWhiteSpace(oneRow[1]).Contains(search_nospace)))
                    {
                        int currentIdx = this.viewList.IndexOf(oneRow);


                        //Third, fill cells.
                        List<int> viewIndex = new List<int> {this.viewColumnName.IndexOf("책이름"), 
                                                            this.viewColumnName.IndexOf("저자"),
                                                            this.viewColumnName.IndexOf("정가"),
                                                            this.viewColumnName.IndexOf("판매가"),
                                                            this.viewColumnName.IndexOf("남은권수"),
                                                            };

                        for (int i = 0; i < viewIndex.Count; i++)
                        {
                            int idx = viewIndex[i];
                            dataGridView1.Rows[rowIdx].Cells[i].Value = this.viewList[currentIdx][idx];
                        }

                        rowIdx++;
                    }
                }
            }
        }

        private void Update_dataGridView2()
        {
            int IsSoldIdx = this.columnName.IndexOf("판매여부");
            int WhoBuyIdx = this.columnName.IndexOf("구매자");

            List<string> displayColumnList = new List<string> { "책이름", "저자", "판매가", "판매시간", "결제방법", "구매자"};

            // Add column headers.
            foreach (string colname in displayColumnList)
            {
                DataGridViewColumn column = new DataGridViewTextBoxColumn();
                column.Name = colname;
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                column.SortMode = DataGridViewColumnSortMode.Automatic;
                dataGridView2.Columns.Add(column);
            }

            // Add contents (each row) if the current row has IsSold=="O".

            // First, Count.
            string search_nospace = this.RemoveAllWhiteSpace(this.search_word);
            int properRowNum = this.dataList.Count(item => (item[IsSoldIdx] == "O") && 
                    (this.RemoveAllWhiteSpace(item[0]).Contains(search_nospace) || this.RemoveAllWhiteSpace(item[1]).Contains(search_nospace) || this.RemoveAllWhiteSpace(item[WhoBuyIdx]).Contains(search_nospace)));

            if (properRowNum > 0)
            {
                // Second, Add rows.
                dataGridView2.RowCount = properRowNum;

                int rowIdx = 0;
                foreach (List<string> oneRow in this.dataList)
                {
                    if (oneRow[IsSoldIdx].Contains("O") && (this.RemoveAllWhiteSpace(oneRow[0]).Contains(search_nospace) || this.RemoveAllWhiteSpace(oneRow[1]).Contains(search_nospace) || this.RemoveAllWhiteSpace(oneRow[WhoBuyIdx]).Contains(search_nospace)))
                    {
                        int currentIdx = this.dataList.IndexOf(oneRow);

                        //Third, fill cells.
                        List<int> viewIndex = new List<int> {this.columnName.IndexOf("책이름"),
                                                            this.columnName.IndexOf("저자"),
                                                            this.columnName.IndexOf("판매가"),
                                                            this.columnName.IndexOf("판매시간"),
                                                            this.columnName.IndexOf("결제방법"),
                                                            this.columnName.IndexOf("구매자"),
                                                            };

                        //Third, fill cells.
                        for (int i = 0; i < viewIndex.Count; i++)
                        {
                            dataGridView2.Rows[rowIdx].Cells[i].Value = this.dataList[currentIdx][viewIndex[i]];
                        }
                        rowIdx++;
                    }
                }
            }
        }

        private void Update_dataGridViews()
        {
            this.dataGridView1.Columns.Clear();
            this.dataGridView1.Rows.Clear();
            this.dataGridView2.Columns.Clear();
            this.dataGridView2.Rows.Clear();
            this.Update_dataGridView1();
            this.Update_dataGridView2();
        }

        private void FormTree_Load()
        {
            string userName = Environment.UserName;
            ListDirectory(this.treeView1, @"C:\Users\" + userName + @"\Desktop\");
        }

        private void ListDirectory(TreeView treeView, string path)
        {
            
            // If there is no nodes, load treeview
            if (treeView.Nodes.Count == 0)
            {
                treeView.Nodes.Clear();
                var rootDirectoryInfo = new DirectoryInfo(path);
                treeView.Nodes.Add(CreateDirectoryNode(rootDirectoryInfo));
            }
        }

        private static TreeNode CreateDirectoryNode(DirectoryInfo directoryInfo)
        {
            var directoryNode = new TreeNode(directoryInfo.Name);

            foreach (var directory in directoryInfo.GetDirectories())
                directoryNode.Nodes.Add(CreateDirectoryNode(directory));
            foreach (var file in directoryInfo.GetFiles())
                directoryNode.Nodes.Add(new TreeNode(file.Name));
            return directoryNode;
        }

        void Panel4_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void Panel4_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            // how to restrict only one drop file?


            // get dropped file path and show on the textbox.
            foreach (string file in files)
                this.textBox2.Text = file;

            // Image shows
            this.ImageShowNextToDiretoryTree();

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // search function.
            this.search_word = this.textBox1.Text;
            this.Update_dataGridViews();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_2(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedIndex = 0;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            this.UpdateDailyChart();
            this.UpdateTotalLabel();
            this.UpdateRatios();
            this.tabControl1.SelectedIndex = 1;
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedIndex = 2;
            // Load file tree
            FormTree_Load();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (this.data_file_path != null)
            {
                this.Save_CurrentExcel();
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.textBox2.Clear();
        }

        private void ImageShowNextToDiretoryTree()
        {
            // Show the appropriate image when user select one file of treeView.
            if (this.textBox2.Text.EndsWith(".xlsx"))
            {
                string imagefilename = Path.Combine(Environment.CurrentDirectory, @"images\fileload_possible.png");
                this.panel4.BackgroundImage = Image.FromFile(imagefilename);
            }

            else
            {
                string imagefilename = Path.Combine(Environment.CurrentDirectory, @"images\fileload_impossible.png");
                this.panel4.BackgroundImage = Image.FromFile(imagefilename);
            }

            this.panel4.BackgroundImageLayout = ImageLayout.Center;
            this.panel4.BackgroundImageLayout = ImageLayout.Zoom;
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            string userName = Environment.UserName;
            this.textBox2.Text = @"C:\Users\" + userName + @"\" + this.treeView1.SelectedNode.FullPath;
            this.ImageShowNextToDiretoryTree();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            // If one excel file is selected,
            if (this.textBox2.Text != "" && this.textBox2.Text.EndsWith(".xlsx"))
            {
                string userName = Environment.UserName;
                this.raw_file_path = this.textBox2.Text;
                this.OpenExcelFile();
                this.tabControl1.SelectedIndex = 0;
                this.Update_dataGridViews();

                // show message on the statusstrip
                DateTime currentTime = DateTime.Now;
                this.toolStripStatusLabel1.Text = "새로운 엑셀 파일을 불러왔습니다 (" + currentTime.ToString("yyyy/MM/dd HH:mm:ss") + ").";
            }
        }


        // Load file. Raise ErrorMessageBox if the file does not exist.
        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.SelectedRows.Count != 0)
            {
                if (this.dataGridView1.SelectedRows.Count > 0 && this.dataGridView1.SelectedRows[0].Cells[0].Value != null)
                {
                    BookInfoWindow BookinfoWindow = new BookInfoWindow(this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString());
                    BookinfoWindow.Show();
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (this.dataGridView2.SelectedRows.Count != 0)
            {
                if (this.dataGridView2.SelectedRows.Count > 0 && this.dataGridView2.SelectedRows[0].Cells[0].Value != null)
                {
                    int WhenSoldIdx = this.columnName.IndexOf("판매시간");
                    int WhenSoldCell = this.dataGridView2.Columns["판매시간"].Index;

                    string selectedBook = this.dataGridView2.SelectedRows[0].Cells[0].Value.ToString();
                    string selectedSoldTime = this.dataGridView2.SelectedRows[0].Cells[WhenSoldCell].Value.ToString();

                    // Find the last book whose name is 'selectedBook' in the dataList.
                    List<string> lastSelectedBookInList = this.dataList.FindLast(
                        delegate (List<string> oneElement)
                        {
                            return oneElement[0] == selectedBook && oneElement[WhenSoldIdx].Equals(selectedSoldTime);
                        });

                    SellInfoWindow SellinfoWindow = new SellInfoWindow(this.columnName, lastSelectedBookInList);
                    SellinfoWindow.Show();
                }
            }
        }
        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox2.Checked && this.checkBox1.Checked)
            {
                this.checkBox1.Checked = true;
                this.checkBox2.Checked = false;
            }
        }

        private void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox1.Checked && this.checkBox2.Checked)
            {
                this.checkBox1.Checked = false;
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.SelectedRows.Count != 0)
            {
                if (this.dataGridView1.SelectedRows.Count > 0 && this.dataGridView1.SelectedRows[0].Cells[0].Value != null)
                {
                    if (this.checkBox1.Checked || this.checkBox2.Checked)
                    {
                        // new window
                        WhoWillBuyWindow whowillbuy_window = new WhoWillBuyWindow();
                        //whowillbuy_window.Show();

                        if (whowillbuy_window.ShowDialog() == DialogResult.OK)
                        {
                            string who = whowillbuy_window.who;
                            string optional = whowillbuy_window.optional;

                            this.SellBook(who: who, optional: optional);
                        }
                    }

                    else if (!this.checkBox1.Checked && !this.checkBox2.Checked)
                    {
                        MessageBox.Show("결제방법 (현금 or 계좌이체)을 선택해주세요!");
                    }
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (this.dataGridView2.SelectedRows.Count != 0)
            {
                if (this.dataGridView2.SelectedRows.Count > 0 && this.dataGridView2.SelectedRows[0].Cells[0].Value != null)
                {
                    this.CancleBook();
                }
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            this.CancleLastUpdate();
        }

        private void SaveMemo()
        {
            StreamWriter writer_;
            string SourceDir = Path.Combine(Environment.CurrentDirectory, @"source");
            string MemoPath = Path.Combine(SourceDir, "memo.txt");
            string CurrentText = this.textBox3.Text;

            writer_ = File.CreateText(MemoPath);
            writer_.Write(CurrentText);
            writer_.Close();
        }

        private void LoadMemo()
        {
            try
            {
                string SourceDir = Path.Combine(Environment.CurrentDirectory, @"source");
                string MemoPath = Path.Combine(SourceDir, "memo.txt");
                string LoadTest = File.ReadAllText(MemoPath);
                this.textBox3.Text = LoadTest;
            }
            catch
            {
                return;
            }

        }


        private void button16_Click(object sender, EventArgs e)
        {
            this.SaveMemo();
        }


        private void InitializeDailyChart()
        {
            System.Windows.Forms.DataVisualization.Charting.Series series = this.chart1.Series[0];
            var hours = Enumerable.Range(00, 24).Select(i => (DateTime.MinValue.AddHours(i)).ToString("HH:mm:ss"));

            foreach (string h in hours)
            {
                series.Points.AddXY(h, 0);
            }

            //this.chart1.ChartAreas[0].AxisX.Maximum = 24;
            //this.chart1.ChartAreas[0].AxisX.Minimum = 0;

            this.UpdateDailyChart();
        }


        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            this.UpdateDailyChart();
        }

        private void UpdateDailyStatistic(int cashtotal, int cashnum, int acctotal, int accnum)
        {
            int total_num = cashnum + accnum;

            if (total_num == 0)
            {
                this.label_statistic_account.Text = "";
                this.label_statistic_cash.Text = "";
                this.label_statistic_soldnum.Text = "";
                this.label_statistic_soldnumperhour.Text = "";
            }
            
            else if (total_num > 0)
            {
                double num_per_hour = total_num / 24.00;

                this.label_statistic_account.Text = Convert.ToString(acctotal)+"원 / "+Convert.ToString(accnum)+"권";
                this.label_statistic_cash.Text = Convert.ToString(cashtotal) + "원 / " + Convert.ToString(cashnum) + "권";
                this.label_statistic_soldnum.Text = Convert.ToString(total_num);
                this.label_statistic_soldnumperhour.Text = Convert.ToString(num_per_hour);
            }
        }

        private void UpdateDailyChart()
        {
            int WhenSoldIdx = this.columnName.IndexOf("판매시간");
            int HowSoldIdx = this.columnName.IndexOf("결제방법");
            int PriceIdx = this.columnName.IndexOf("판매가");
            char TimeSpliter = ':';

            System.Windows.Forms.DataVisualization.Charting.Series series = this.chart1.Series[0];
            var hours = Enumerable.Range(00, 24).Select(i => (DateTime.MinValue.AddHours(i)).ToString("HH:mm:ss"));

            List<string> HourStringList = new List<string>();
            foreach (string h in hours)
            {
                HourStringList.Add(h);
            }

            string selectedDay = this.dateTimePicker1.Value.ToString("yyyy/MM/dd");
            List<double> CountBooks = new List<double>(new double[24]);

            int todaycashnum = 0;
            int todaycashtotal = 0;
            int todayaccnum = 0;
            int todayacctotal = 0;

            foreach (List<string> oneBook in this.dataList)
            {
                if (oneBook[WhenSoldIdx].Split(null)[0]==selectedDay)
                {
                    try
                    {
                        int time = Convert.ToInt32(oneBook[WhenSoldIdx].Split(null)[1].Split(TimeSpliter)[0]);
                        CountBooks[time]++;

                        if (oneBook[HowSoldIdx]=="현금")
                        {
                            todaycashtotal += Convert.ToInt32(oneBook[PriceIdx]);
                            todaycashnum++;
                        }
                        else if (oneBook[HowSoldIdx]== "계좌이체")
                        {
                            todayacctotal += Convert.ToInt32(oneBook[PriceIdx]);
                            todayaccnum++;
                        }
                    }
                    catch (Exception ex) 
                    {
                        MessageBox.Show(ex.Message + " !");
                    }
                }
            }

            try
            {
                // Update chart
                series.Points.Clear();

                int hidx = 0;
                foreach (string h in hours)
                {
                    series.Points.AddXY(h, CountBooks[hidx]);
                    series.Points[series.Points.Count - 1].ToolTip = "time: " + h +"~"+Convert.ToString(Convert.ToInt32(h.Split(TimeSpliter)[0])+1)+":00:00";
                    hidx++;
                }

                // Update infos
                this.UpdateDailyStatistic(todaycashtotal, todaycashnum, todayacctotal, todayaccnum);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " !");
            }
        }
    }   
}

