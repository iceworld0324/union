using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace UnionFundsCalculator
{
    public partial class MainWindow : Window
    {
        private const string timeFormat_ = "yyyy年MM月"; // 时间的显示格式
        private const string numberFormat_ = "0.00"; // 生成表格的数字格式
        private const string font_ = "宋体"; // 生成表格的字体
        private const int startYear_ = 2016; // 年份下拉菜单的开始年

        private DatabaseInterface database_; // 数据库接口
        private string filenameFunds_; // 经费信息表格的文件名
        private Dictionary<string, List<string>> categories_; // 所有分类，如（工会->所有工会名称）
        private DataTable settings_; // 系统设置的内存对象
        private string companyId_; // 用户输入的公司名
        private Dictionary<string, int> selectedCategory_; // 用户选择的分类，如（工会->0）意味着用户选择了第一个工会

        public MainWindow()
        {
            InitializeComponent();
            InitializeComboBox(comboBoxInputYear, comboBoxInputMonth);
            InitializeComboBox(comboBoxOutputYear, comboBoxOutputMonth);
            database_ = new DatabaseInterface();
            categories_ = new Dictionary<string, List<string>>
            {
                { "TaxAuthority", new List<string>() },
                { "Union", new List<string>() },
                { "System", new List<string>() },
                { "Industry", new List<string>() },
            };
            UpdateCategories();
            InitializeSettings();
        }

        // 初始化选择年月的下拉菜单
        private void InitializeComboBox(ComboBox comboBoxYear, ComboBox comboBoxMonth)
        {
            DateTime today = DateTime.Today;
            for (int year = startYear_; year <= today.Year; year++)
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Content = year;
                comboBoxYear.Items.Add(item);
            }
            comboBoxYear.SelectedIndex = today.Year - startYear_;
            for (int month = 1; month <= 12; month++)
            {
                ComboBoxItem item = new ComboBoxItem();
                item.Content = month;
                comboBoxMonth.Items.Add(item);
            }
            comboBoxMonth.SelectedIndex = today.Month - 1;
        }

        // 根据数据库中的公司信息来填写categories_
        private void UpdateCategories()
        {
            foreach (KeyValuePair<string, List<string>> entry in categories_)
            {
                entry.Value.Clear();
                string sql = "SELECT DISTINCT [" + entry.Key + "] FROM CompanyInfo WHERE [" + entry.Key + "] IS NOT NULL";
                using (var cmd = new SqlCeCommand(sql, database_.GetConnection()))
                {
                    SqlCeDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        entry.Value.Add(reader.GetString(0));
                    }
                    reader.Close();
                }
            }
        }

        // 根据数据库中的设置信息来填写settings_
        private void InitializeSettings()
        {
            settings_ = new DataTable();
            using (var cmd = new SqlCeCommand("SELECT * FROM Settings", database_.GetConnection()))
            {
                SqlCeDataAdapter adapter = new SqlCeDataAdapter(cmd);
                adapter.Fill(settings_);
                dataGridSettings.ItemsSource = settings_.DefaultView;
            }
        }

        // 显示数据库中已存储的数据
        private void Display_Data_Button_Click(object sender, RoutedEventArgs e)
        {
            string message = "数据库现有以下数据\n";
            using (var cmd = new SqlCeCommand("SELECT COUNT(*) FROM CompanyInfo", database_.GetConnection()))
            {
                object count = cmd.ExecuteScalar();
                message += "公司信息" + count + "条\n";
            }
            using (var cmd = new SqlCeCommand("SELECT DISTINCT Time FROM Funds", database_.GetConnection()))
            {
                message += "经费信息：";
                SqlCeDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    message += (reader.GetDateTime(0).ToString(timeFormat_) + " ");
                }
                reader.Close();
            }
            textBoxMessage.Text = message;
        }

        /* 把公司信息表格中的内容写入数据库
        private void Company_Button_Click(object sender, RoutedEventArgs e)
        {
            if (filenameCompanyInfo_ == null)
            {
                textBoxMessage.Text = "请先用浏览键选择一个包含公司信息的Excel文件";
                return;
            }

            ExcelInterface excel = new ExcelInterface(filenameCompanyInfo_, 'r');
            Excel.Range range = excel.GetWorksheet().UsedRange;
            int row = range.Rows.Count;
            SqlCeConnection connection = database_.GetConnection();
            using (var cmd = new SqlCeCommand("DELETE CompanyInfo", connection))
            {
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCeCommand("INSERT INTO CompanyInfo VALUES (?, ?, ?, ?, ?, ?)", connection))
            {
                cmd.Parameters.Add("@id", SqlDbType.NVarChar);
                cmd.Parameters.Add("@name", SqlDbType.NVarChar);
                cmd.Parameters.Add("@tax", SqlDbType.NVarChar);
                cmd.Parameters.Add("@union", SqlDbType.NVarChar);
                cmd.Parameters.Add("@system", SqlDbType.NVarChar);
                cmd.Parameters.Add("@industry", SqlDbType.NVarChar);

                SqlCeTransaction transaction = connection.BeginTransaction();
                cmd.Transaction = transaction;
                try
                {
                    for (int rCnt = 2; rCnt <= row; rCnt++)
                    {
                        for (int cCnt = 1; cCnt <= 6; cCnt++)
                        {
                            // 如果excel cell为空，则把dbnull.value写入数据库
                            cmd.Parameters[cCnt - 1].Value =
                                (range.Cells[rCnt, cCnt] as Excel.Range).Value ?? DBNull.Value;
                        }
                        cmd.ExecuteNonQuery();
                    }
                    transaction.Commit();
                    UpdateCategories();
                    textBoxMessage.Text = "更新成功，数据库现有" + (row - 1) + "条公司信息";
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    textBoxMessage.Text = "更新失败，提供的Excel文件存在问题。详细信息：" + ex.Message;
                }
            }
            excel.Close();
        }*/

        // 把经费信息表格中的内容写入数据库
        private void Add_Funds_Button_Click(object sender, RoutedEventArgs e)
        {
            if (filenameFunds_ == null)
            {
                textBoxMessage.Text = "请先用浏览键选择一个包含经费信息的Excel文件";
                return;
            }

            ExcelInterface excel = new ExcelInterface(filenameFunds_, 'r');
            Excel.Range range = GetSortedRange(excel);
            int row = range.Rows.Count;
            SqlCeConnection connection = database_.GetConnection();
            DateTime time = GetTimeFromComboBox(comboBoxInputYear, comboBoxInputMonth);

            using (var cmd = new SqlCeCommand("DELETE FROM Funds WHERE Time = @time", connection))
            {
                cmd.Parameters.AddWithValue("@time", time);
                cmd.ExecuteNonQuery();
            }

            using (var cmd = new SqlCeCommand("INSERT INTO Funds VALUES (?, ?, ?)", connection))
            {
                cmd.Parameters.Add("@id", SqlDbType.NVarChar);
                cmd.Parameters.Add("@time", SqlDbType.DateTime);
                cmd.Parameters.Add("@received", SqlDbType.Float);
                SqlCeTransaction transaction = connection.BeginTransaction();
                cmd.Transaction = transaction;
                try
                {
                    float totalReceived = 0; // 相邻行如果是同一公司则把几个经费数字加起来
                    for (int rCnt = 2; rCnt <= row; rCnt++)
                    {
                        string id = (string)(range.Cells[rCnt, 1] as Excel.Range).Value;
                        totalReceived += (float)(range.Cells[rCnt, 5] as Excel.Range).Value;
                        if (rCnt == row || id != (string)(range.Cells[rCnt + 1, 1] as Excel.Range).Value)
                        {
                            cmd.Parameters[0].Value = id;
                            cmd.Parameters[1].Value = time;
                            cmd.Parameters[2].Value = totalReceived;
                            cmd.ExecuteNonQuery();
                            totalReceived = 0;
                        }
                    }
                    transaction.Commit();
                    textBoxMessage.Text = "成功添加" + time.ToString(timeFormat_) + (row - 1) + "条经费信息";
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    textBoxMessage.Text = "添加失败，提供的Excel文件存在问题，请确保文件最后的合计行之后没有多余的内容。详细信息：" + ex.Message;
                }
            }
            excel.Close();
        }

        private Excel.Range GetSortedRange(ExcelInterface excel)
        {
            Excel.Range range = excel.GetWorksheet().UsedRange;
            range.Rows[range.Rows.Count].Delete(); // 删除最后的合计行
            range.Sort(range.Columns[1], Header: Excel.XlYesNoGuess.xlYes); // 对除标题行外的其他行排序
            return range;
        }

        // 根据经费信息表格自动更新公司信息
        private void Auto_Update_Company_Button_Click(object sender, RoutedEventArgs e)
        {
            if (filenameFunds_ == null)
            {
                textBoxMessage.Text = "请先用浏览键选择一个包含经费信息的Excel文件";
                return;
            }

            ExcelInterface excel = new ExcelInterface(filenameFunds_, 'r');
            Excel.Range range = GetSortedRange(excel);
            int row = range.Rows.Count;
            var toInsert = new List<Company>();
            var toUpdate = new List<Company>();

            using (var cmd = new SqlCeCommand("SELECT TaxAuthority FROM CompanyInfo WHERE CompanyId = @id", database_.GetConnection()))
            {
                cmd.Parameters.Add("@id", SqlDbType.NVarChar);
                for (int rCnt = 2; rCnt <= row; rCnt++)
                {
                    var item = new Company();
                    item.id = (string)(range.Cells[rCnt, 1] as Excel.Range).Value;
                    item.name = (string)(range.Cells[rCnt, 2] as Excel.Range).Value;
                    item.tax = (string)(range.Cells[rCnt, 6] as Excel.Range).Value;
                    if (rCnt < row && item.id == (string)(range.Cells[rCnt + 1, 1] as Excel.Range).Value)
                    {
                        continue;
                    }
                    DecideInsertOrUpdate(cmd, item, toInsert, toUpdate);
                }
            }

            excel.Close();
            InsertNewCompanies(toInsert);
            UpdateExistingCompanies(toUpdate);
            UpdateCategories();
        }

        private void DecideInsertOrUpdate(SqlCeCommand cmd, Company com, List<Company> toInsert, List<Company> toUpdate)
        {
            cmd.Parameters[0].Value = com.id;
            SqlCeDataReader reader = cmd.ExecuteResultSet(ResultSetOptions.Scrollable);
            if (reader.HasRows)
            {
                reader.Read();
                if (reader.GetString(0) != com.tax)
                {
                    toUpdate.Add(com); // 公司已存在但税收机关不正确
                }
            }
            else
            {
                toInsert.Add(com); // 公司不存在
            }
            reader.Close();
        }

        private void InsertNewCompanies(List<Company> toInsert)
        {
            SqlCeConnection connection = database_.GetConnection();
            using (var cmd = new SqlCeCommand("INSERT INTO CompanyInfo VALUES (?, ?, ?, ?, ?, ?)", connection))
            {
                cmd.Parameters.Add("@id", SqlDbType.NVarChar);
                cmd.Parameters.Add("@name", SqlDbType.NVarChar);
                cmd.Parameters.Add("@tax", SqlDbType.NVarChar);
                cmd.Parameters.AddWithValue("@union", DBNull.Value);
                cmd.Parameters.AddWithValue("@system", DBNull.Value);
                cmd.Parameters.AddWithValue("@industry", DBNull.Value);

                SqlCeTransaction transaction = connection.BeginTransaction();
                cmd.Transaction = transaction;
                try
                {
                    foreach (Company item in toInsert)
                    {
                        cmd.Parameters[0].Value = item.id;
                        cmd.Parameters[1].Value = item.name;
                        cmd.Parameters[2].Value = item.tax;
                        cmd.ExecuteNonQuery();
                    }
                    transaction.Commit();
                    textBoxMessage.Text = "成功添加" + toInsert.Count + "条新公司信息";
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    textBoxMessage.Text = "添加新公司信息失败。详细信息：" + ex.Message;
                }
            }
        }

        private void UpdateExistingCompanies(List<Company> toUpdate)
        {
            SqlCeConnection connection = database_.GetConnection();
            using (var cmd = new SqlCeCommand("UPDATE CompanyInfo SET TaxAuthority = @tax WHERE CompanyID = @id", connection))
            {
                cmd.Parameters.Add("@tax", SqlDbType.NVarChar);
                cmd.Parameters.Add("@id", SqlDbType.NVarChar);
             
                SqlCeTransaction transaction = connection.BeginTransaction();
                cmd.Transaction = transaction;
                try
                {
                    foreach (Company item in toUpdate)
                    {
                        cmd.Parameters[0].Value = item.tax;
                        cmd.Parameters[1].Value = item.id;
                        cmd.ExecuteNonQuery();
                    }
                    transaction.Commit();
                    textBoxMessage.AppendText("\n成功更新" + toUpdate.Count + "条已有公司信息");
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    textBoxMessage.AppendText("\n更新已有公司信息失败。详细信息：" + ex.Message);
                }
            }
        }

        // 通过一个新窗口手动更改公司信息
        private void Manual_Update_Company_Button_Click(object sender, RoutedEventArgs e)
        {
            UpdateCompanyWindow window = new UpdateCompanyWindow(database_);
            window.ShowDialog();
            if (window.DialogResult == true)
            {
                UpdateCategories();
                textBoxMessage.Text = "公司信息已保存";
            }
            else
            {
                textBoxMessage.Text = "公司信息的更改已取消";
            }
        }

        // 浏览键
        private void Browse_Funds_Button_Click(object sender, RoutedEventArgs e)
        {
            string filename = RunBrowseDialog();
            filenameFunds_ = filename;
            textBoxFunds.Text = System.IO.Path.GetFileName(filename);
        }

        private string RunBrowseDialog()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "Excel文件 (*.xls, *.xlsx)|*.xls;*.xlsx";
            Nullable<bool> result = dlg.ShowDialog();
            return result == true ? dlg.FileName : null;
        }

        // 从下拉菜单读取用户选择的年月
        private DateTime GetTimeFromComboBox(ComboBox comboBoxYear, ComboBox comboBoxMonth)
        {
            ComboBoxItem itemYear = (ComboBoxItem)comboBoxYear.SelectedItem;
            ComboBoxItem itemMonth = (ComboBoxItem)comboBoxMonth.SelectedItem;
            return new DateTime((int)itemYear.Content, (int)itemMonth.Content, 1);
        }

        // 生成一个月的所有报表
        private void Generate_Report_Button_Click(object sender, RoutedEventArgs e)
        {
            DateTime time = GetTimeFromComboBox(comboBoxOutputYear, comboBoxOutputMonth);
            using (var cmd = new SqlCeCommand("SELECT TOP 1 CompanyId FROM Funds WHERE Time = @time", database_.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@time", time);
                SqlCeDataReader reader = cmd.ExecuteReader();
                if (reader.Read()) // 该月有至少一条经费信息
                {
                    GroupByIndustry(time);
                    GroupByTaxAuthority(time);
                    GroupByUnion(time);
                    textBoxMessage.Text = "生成完毕，报表已存至目录：\"" + Directory.GetCurrentDirectory() + "\"";
                }
                else // 该月没有经费信息
                {
                    textBoxMessage.Text = "生成失败，数据库中没有" + time.ToString(timeFormat_) + "的数据";
                }
            }
        }

        // 生成总工会所属公司按产业分组的报表
        private void GroupByIndustry(DateTime time)
        {
            ExcelInterface excel = new ExcelInterface(null, 'w');
            FillCells(excel.GetWorksheet().Cells, time, "三门峡市总工会", "Industry");
            StyleCells(excel.GetWorksheet().Cells);
            string filename = Path.Combine(Directory.GetCurrentDirectory(),
                    time.ToString(timeFormat_) + "工会经费报表（产业）.xls");
            excel.Save(filename);    
            excel.Close();
        }

        // 生成总工会所属公司按税务局分组的报表
        private void GroupByTaxAuthority(DateTime time)
        {
            ExcelInterface excel = new ExcelInterface(null, 'w');
            FillCells(excel.GetWorksheet().Cells, time, "三门峡市总工会", "TaxAuthority");
            StyleCells(excel.GetWorksheet().Cells);
            string filename = Path.Combine(Directory.GetCurrentDirectory(),
                    time.ToString(timeFormat_) + "工会经费征收明细.xls");
            excel.Save(filename);
            excel.Close();
        }

        // 填写cells的内容，具体为time时间、union工会下的经费信息，按groupBy分组
        private void FillCells(Excel.Range cells, DateTime time, string union, string groupBy)
        {
            cells[1, 1] = "纳税人名称";
            cells[1, 2] = "实缴金额";
            string sql = @"SELECT c.CompanyName, f.Received FROM CompanyInfo c INNER JOIN Funds f
                         ON c.CompanyID = f.CompanyID WHERE f.Time = @time AND c."
                         + groupBy + " = @group" + (union != null ? " AND c.[Union] = '" + union + "'" : "");
            using (var cmd = new SqlCeCommand(sql, database_.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@time", time);
                cmd.Parameters.Add("@group", SqlDbType.NVarChar);
                double fileTotal = 0;
                int row = 2;
                foreach (string group in categories_[groupBy])
                {
                    cmd.Parameters[1].Value = group;
                    SqlCeDataReader reader = cmd.ExecuteReader();
                    double groupTotal = 0;
                    while (reader.Read())
                    {
                        cells[row, 1] = reader.GetString(0);
                        cells[row, 2] = Math.Round(reader.GetDouble(1), 2);
                        groupTotal += cells[row, 2].Value;
                        ++row;
                    }
                    reader.Close();
                    cells.Rows[row].Font.Bold = true;
                    cells[row, 1] = group;
                    cells[row, 2] = groupTotal;
                    fileTotal += groupTotal;
                    groupTotal = 0;
                    ++row;
                }
                cells.Rows[row].Font.Bold = true;
                cells[row, 1] = "合计";
                cells[row, 2] = fileTotal;
            }
        }

        // 设置cells的格式
        private void StyleCells(Excel.Range cells)
        {
            cells.NumberFormat = numberFormat_;
            cells.Font.Name = font_;
            cells.Font.Size = 9;
            cells.Columns.AutoFit();
        }

        // 生成全部公司按工会分组的报表
        private void GroupByUnion(DateTime time)
        {            
            ExcelInterface excel = new ExcelInterface(null, 'w');
            GroupByUnionFillCells(excel.GetWorksheet().Cells, time);
            GroupByUnionStyleCells(excel.GetWorksheet().Cells);
            string filename = Path.Combine(Directory.GetCurrentDirectory(),
                time.ToString(timeFormat_) + "各县（市）区总工会经费收解返计算表.xls");
            excel.Save(filename);
            excel.Close();
        }

        private void GroupByUnionFillCells(Excel.Range cells, DateTime time)
        {
            // 填写前四行
            cells[2, 1] = "各县（市）区总工会经费收解返计算表";
            string[] labels = new string[] { "序号", "单位名称", "地税机关代收金额", "应上解市总经费", "实际上解", "实际返拨经费" };
            cells.Range[cells[4, 1], cells[4, 6]].Value = labels;
            string sql = @"SELECT SUM(f.Received) FROM CompanyInfo c INNER JOIN Funds f
                         ON c.CompanyId = f.CompanyId WHERE f.Time = @time AND c.[Union] = @union";
            int lastRow = 6;       
            // 填写中间部分
            using (var cmd = new SqlCeCommand(sql, database_.GetConnection()))
            {
                cmd.Parameters.AddWithValue("@time", time);
                cmd.Parameters.Add("@union", SqlDbType.NVarChar);
                string ratio = GetSetting("应上解市总经费的比例");
                foreach (string union in categories_["Union"])
                {
                    cmd.Parameters[1].Value = union;
                    int row;
                    if (union == "三门峡市总工会") row = 5; // 这两个工会必须在其他工会前面
                    else if (union == "开发区工会办事处") row = 6;
                    else row = ++lastRow;
                    cells[row, 2] = union;
                    cells[row, 3] = Math.Round((double)cmd.ExecuteScalar(), 2);
                    if (row >= 7)
                    {
                        cells[row, 4].Formula = "=C" + row + "*" + ratio;
                        cells[row, 6].Formula = "=C" + row + "-E" + row;
                    }
                }
            }
            // 填写后两行
            cells[lastRow + 1, 2] = "2-" + (lastRow - 4) + "小计";
            cells[lastRow + 2, 2] = "合计";
            for (int col = 3; col <= 6; col++)
            {
                char charCol = (char)(col + 64);
                cells[lastRow + 1, col].Formula = "=SUM(" + charCol + "6:" + charCol + lastRow + ")";
                cells[lastRow + 2, col].Formula = "=" + charCol + "5" + "+" + charCol + (lastRow + 1);
            }
            // 填写第一列
            for (int row = 5; row <= lastRow + 2; row++)
            {
                cells[row, 1] = row - 4;
            }       
        }

        private void GroupByUnionStyleCells(Excel.Range cells)
        {
            cells.Range[cells[2, 1], cells[2, 6]].Merge();
            cells[2, 1].Font.Bold = true;
            cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            int lastRow = cells.Find("*", SearchOrder:Excel.XlSearchOrder.xlByRows,
                SearchDirection:Excel.XlSearchDirection.xlPrevious).Row;
            cells.Range[cells[4, 6], cells[lastRow - 1, 6]].Font.Bold = true;
            cells.Range[cells[4, 1], cells[lastRow - 1, 6]].RowHeight = 56.25;
            cells.Range[cells[4, 1], cells[lastRow, 6]].Borders.Weight = 2d;
            cells.Range[cells[5, 3], cells[lastRow, 6]].NumberFormat = numberFormat_;
            cells.Font.Name = font_;
            cells.Font.Size = 12;
            cells.Columns.AutoFit();
            cells.Columns[5].ColumnWidth = cells.Columns[4].ColumnWidth;
        }

        // 激活选择公司窗口，并在窗口关闭后获取选择的结果
        private void Choose_Company_Button_Click(object sender, RoutedEventArgs e)
        {
            ChooseCompanyWindow window = new ChooseCompanyWindow(categories_, database_);
            if (window.ShowDialog() == true)
            {
                if (window.radioButtonName.IsChecked == true)
                {
                    companyId_ = window.companyId;
                    textBoxMessage.Text = "已选择公司：" + window.textBoxCompanyName.Text;
                    selectedCategory_ = null;
                }
                else if (window.radioButtonCategory.IsChecked == true)
                {
                    selectedCategory_ = window.selectedCategory;
                    textBoxMessage.Text = "已选择分类：" + GetMessageFromSelectedCategory(categories_, selectedCategory_);
                    companyId_ = null;
                }
            }
        }

        private string GetMessageFromSelectedCategory(Dictionary<string, List<string>> categories, Dictionary<string, int> selectedCategory)
        {
            List<string> words = new List<string>();
            foreach (KeyValuePair<string, int> entry in selectedCategory)
            {
                if (entry.Value != -1)
                {
                    words.Add(categories[entry.Key][entry.Value]);
                }
            }
            return (words.Count > 0) ? string.Join("，", words) : "全部";
        }

        // 生成按时间比较的结果
        private void Generate_Comparison_Button_Click(object sender, RoutedEventArgs e)
        {
            int period;
            if (!int.TryParse(textBoxPeriod.Text, out period) || period < 0)
            {
                textBoxMessage.Text = "请在文本框中输入一个非负整数";
                return;
            }
            if (companyId_ == null && selectedCategory_ == null)
            {
                textBoxMessage.Text = "请先用选择公司键选择一个或一类公司";
                return;
            }

            string statement = (companyId_ != null) ?
                "SELECT Received FROM Funds WHERE Time = @time AND CompanyId = '" + companyId_ + "'":
                "SELECT SUM(f.Received) FROM CompanyInfo c INNER JOIN Funds f ON c.CompanyId = f.CompanyId WHERE f.Time = @time" +
                GetStatementFromSelectedCategory(categories_, selectedCategory_);
            DateTime time = GetTimeFromComboBox(comboBoxOutputYear, comboBoxOutputMonth);
            var list = new List<ComparisonResult>();
            using (var cmd = new SqlCeCommand(statement, database_.GetConnection()))
            {
                cmd.Parameters.Add("@time", SqlDbType.DateTime);
                for (int i = 0; i <= period; i++)
                {
                    cmd.Parameters[0].Value = time;
                    object received = cmd.ExecuteScalar();
                    var item = new ComparisonResult();
                    // 当select结果是0行时，received返回null但是sum(f.received)返回dbnull.value
                    // math.round只接受double，而用?语句给nullable double赋值时只接受nullable double
                    item.received = (received != null && received != DBNull.Value) ? (double?)Math.Round((double)received, 2) : null;
                    item.time = time.ToString(timeFormat_);
                    list.Add(item);
                    time = (comboBoxPeriod.SelectedIndex == 0) ? time.AddMonths(-1) : time.AddYears(-1);
                }
            }
            var window = new DataDisplayWindow(list);
            window.ShowDialog();
        }

        private string GetStatementFromSelectedCategory(Dictionary<string, List<string>> categories,
            Dictionary<string, int> selectedCategory)
        {
            string statement = "";
            foreach (KeyValuePair<string, int> entry in selectedCategory)
            {
                if (entry.Value != -1)
                {
                    statement += (" AND c.[" + entry.Key + "]='" + categories[entry.Key][entry.Value] + "'");
                }
            }
            return statement;
        }

        // 保存设置到数据库
        private void Save_Settings_Button_Click(object sender, RoutedEventArgs e)
        {
            using (var cmd = new SqlCeCommand("SELECT * FROM Settings", database_.GetConnection()))
            {
                SqlCeDataAdapter adapter = new SqlCeDataAdapter(cmd);
                SqlCeCommandBuilder builder = new SqlCeCommandBuilder(adapter);
                adapter.UpdateCommand = builder.GetUpdateCommand();
                adapter.Update(settings_);
            }
            settings_.AcceptChanges();
            textBoxMessage.Text = "设置已保存";
        }

        // 从设置标签离开时，丢弃未保存的设置
        private void Tab_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ReferenceEquals(e.OriginalSource, tabOverall) &&
                e.RemovedItems.Count > 0 &&
                (string)((TabItem)e.RemovedItems[0]).Header == "设置")
            {
                settings_.RejectChanges();
            }
        }

        // 获取一个设置的值
        private string GetSetting(string name)
        {
            DataRow[] foundRows = settings_.Select("Name = '" + name + "'");
            return foundRows[0][1].ToString();
        }
    }

    public class ExcelInterface
    {
        private Excel.Application app_;
        private Excel.Workbook workBook_;
        private Excel.Worksheet workSheet_;

        public ExcelInterface(string filename, char mode)
        {
            app_ = new Excel.Application();
            app_.DisplayAlerts = false; // 保存时直接覆盖重名的文件，不会弹出询问用户的窗口
            if (mode == 'r')
            {
                workBook_ = app_.Workbooks.Open(filename, 0, true, 5, "", "", true,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                    "\t", false, false, 0, true, 1, 0);
            }
            else if (mode == 'w')
            {
                workBook_ = app_.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            }
            workSheet_ = (Excel.Worksheet)workBook_.Worksheets.get_Item(1);
        }

        public void Close()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.FinalReleaseComObject(workSheet_);
            workBook_.Close(SaveChanges:false);
            Marshal.FinalReleaseComObject(workBook_);
            app_.Quit();
            Marshal.FinalReleaseComObject(app_);
        }

        public void Save(string filename)
        {
            workBook_.SaveAs(filename);
        }

        public Excel.Worksheet GetWorksheet()
        {
            return workSheet_;
        }
    }

    public class DatabaseInterface
    {
        private SqlCeConnection connection_;

        public DatabaseInterface()
        {
            connection_ = new SqlCeConnection("Data Source=|DataDirectory|\\Database1.sdf");
            connection_.Open();
        }

        ~DatabaseInterface()
        {
            connection_.Close();
        }

        public SqlCeConnection GetConnection()
        {
            return connection_;
        }
    }

    public class ComparisonResult
    {
        [DisplayName("时间")]
        public string time { get; set; }
        [DisplayName("实缴金额")]
        public double? received { get; set; }
    }

    public class Company
    {
        [DisplayName("纳税人识别号")]
        public string id { get; set; }
        [DisplayName("纳税人名称")]
        public string name { get; set; }
        [DisplayName("税务机关")]
        public string tax { get; set; }
    }
}