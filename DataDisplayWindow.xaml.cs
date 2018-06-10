using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace UnionFundsCalculator
{
    public partial class DataDisplayWindow : Window
    {
        // 两个构造函数接受两个不同的list
        public DataDisplayWindow(List<ComparisonResult> list)
        {
            InitializeComponent();
            dataGridDisplay.ItemsSource = list;
            Title = "比较结果";
        }

        public DataDisplayWindow(List<Company> list)
        {
            InitializeComponent();
            dataGridDisplay.ItemsSource = list;
            Title = "新增公司";
        }

        // 生成datagrid时显示正确的列名称
        private void DataGrid_Display_Generating_Column(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            var displayName = GetPropertyDisplayName(e.PropertyDescriptor);
            if (!string.IsNullOrEmpty(displayName))
            {
                e.Column.Header = displayName;
            }

        }

        public static string GetPropertyDisplayName(object descriptor)
        {
            var pd = descriptor as PropertyDescriptor;
            if (pd != null)
            {
                var displayName = pd.Attributes[typeof(DisplayNameAttribute)] as DisplayNameAttribute;
                if (displayName != null && displayName != DisplayNameAttribute.Default)
                {
                    return displayName.DisplayName;
                }
            }
            else
            {
                var pi = descriptor as PropertyInfo;
                if (pi != null)
                {
                    Object[] attributes = pi.GetCustomAttributes(typeof(DisplayNameAttribute), true);
                    for (int i = 0; i < attributes.Length; ++i)
                    {
                        var displayName = attributes[i] as DisplayNameAttribute;
                        if (displayName != null && displayName != DisplayNameAttribute.Default)
                        {
                            return displayName.DisplayName;
                        }
                    }
                }
            }
            return null;
        }

        // 复制datagrid中的数据到剪贴板
        private void Copy_Button_Click(object sender, RoutedEventArgs e)
        {
            dataGridDisplay.SelectAllCells();
            dataGridDisplay.ClipboardCopyMode = DataGridClipboardCopyMode.ExcludeHeader;
            ApplicationCommands.Copy.Execute(null, dataGridDisplay);
            dataGridDisplay.UnselectAllCells();
        }
    }
}