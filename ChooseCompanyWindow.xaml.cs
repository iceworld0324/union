using System.Collections.Generic;
using System.Data.SqlServerCe;
using System.Windows;

namespace UnionFundsCalculator
{
    public partial class ChooseCompanyWindow : Window
    {
        private DatabaseInterface database_; // 数据库接口

        public string companyId { get; set; } // 用户输入的公司名
        public Dictionary<string, int> selectedCategory { get; set; } // 用户选择的分类，如（工会->0）意味着用户选择了第一个工会

        public ChooseCompanyWindow(Dictionary<string, List<string>> categories, DatabaseInterface database)
        {
            InitializeComponent();
            comboBoxTax.ItemsSource = categories["TaxAuthority"];
            comboBoxUnion.ItemsSource = categories["Union"];
            comboBoxSystem.ItemsSource = categories["System"];
            comboBoxIndustry.ItemsSource = categories["Industry"];
            database_ = database;
        }

        // 确认用户的选择并关闭窗口
        private void Confirm_Button_Click(object sender, RoutedEventArgs e)
        {
            if (radioButtonName.IsChecked == true) // 输入公司名
            {
                using (var cmd = new SqlCeCommand("SELECT CompanyId FROM CompanyInfo WHERE CompanyName = @name", database_.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@name", textBoxCompanyName.Text);
                    companyId = (string)cmd.ExecuteScalar();
                }
                if (companyId == null)
                {
                    MessageBox.Show("错误：无法在数据库包含的公司信息中找到输入的公司名称");
                }
                else
                {
                    DialogResult = true;
                }
            }
            else if (radioButtonCategory.IsChecked == true) // 选择分类
            {
                selectedCategory = new Dictionary<string, int>
                {
                    { "TaxAuthority", comboBoxTax.SelectedIndex },
                    { "Union", comboBoxUnion.SelectedIndex },
                    { "System", comboBoxSystem.SelectedIndex },
                    { "Industry", comboBoxIndustry.SelectedIndex },
                };
                DialogResult = true;
            }
        }
    }
}