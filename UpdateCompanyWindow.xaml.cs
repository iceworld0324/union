using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Windows;

namespace UnionFundsCalculator
{
    public partial class UpdateCompanyWindow : Window
    {
        private DatabaseInterface database_; // 数据库接口
        private DataTable company_; // 公司信息的内存对象

        public UpdateCompanyWindow(DatabaseInterface database)
        {
            InitializeComponent();
            database_ = database;
            company_ = new DataTable();
            FillDataGrid();
        }

        private void FillDataGrid()
        {
            using (var cmd = new SqlCeCommand("SELECT * FROM CompanyInfo", database_.GetConnection()))
            {
                SqlCeDataAdapter adapter = new SqlCeDataAdapter(cmd);
                adapter.Fill(company_);
                dataGridCompany.ItemsSource = company_.DefaultView;
                dataGridCompany.Items.SortDescriptions.Add(
                    new SortDescription("Union", ListSortDirection.Ascending));
            }
        }

        private void Save_Button_Click(object sender, RoutedEventArgs e)
        {
            using (var cmd = new SqlCeCommand("SELECT * FROM CompanyInfo", database_.GetConnection()))
            {
                SqlCeDataAdapter adapter = new SqlCeDataAdapter(cmd);
                SqlCeCommandBuilder builder = new SqlCeCommandBuilder(adapter);
                adapter.UpdateCommand = builder.GetUpdateCommand();
                adapter.Update(company_);
            }
            company_.AcceptChanges();
            DialogResult = true;
        }

        private void Cancel_Button_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
