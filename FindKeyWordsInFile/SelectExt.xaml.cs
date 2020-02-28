using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace FindKeyWordsInFile
{
    /// <summary>
    /// SelectExt.xaml 的交互逻辑
    /// </summary>
    public partial class SelectExt : Window
    {
        public string Exts = "";
        public SelectExt(string exts)
        {
            Exts = exts;
            InitializeComponent();
            this.Loaded += load;
        }

        private void load(object sender, RoutedEventArgs e)
        {
            string[] curExtsArry = Exts.Split(';').Select(z => z.Replace("*.","")).ToArray();
            string exts =  ConfigurationManager.AppSettings["exts"].ToString();
            string[] extsArry =  exts.Split(';');
            List<string>  extsList = extsArry.Where(xz => !string.IsNullOrWhiteSpace(xz)).ToList();
            extsList.Insert(0, "全选");
            int rowNum = grid.RowDefinitions.Count;
            int colNum = grid.ColumnDefinitions.Count;
            int x = 0;
            int y = 0;
            int num = 0;
            int index = 0;
            foreach (string item in extsList)
            {
                CheckBox checkBox = new CheckBox();
                index++;
                num++;
                if (index == 1)
                {
                    checkBox.IsChecked = true;
                    checkBox.Checked += SelectAll;
                    checkBox.Unchecked += UnSelectAll;
                }
                if (curExtsArry.Contains(item))
                {
                    checkBox.IsChecked = true;
                }
                checkBox.Content = item;
                grid.Children.Add(checkBox);
                Grid.SetColumn(checkBox, x);
                Grid.SetRow(checkBox, y);
                y++;
                if (num >= rowNum)
                {
                    x++;
                    num = 0;
                    y = 0;
                }
            }
        }

        private void SelectAll(object sender, RoutedEventArgs e)
        {
            foreach (var item in grid.Children)
            {
                if (item is CheckBox)
                {
                    if (((CheckBox)item).Content.ToString() == "全选" )
                    {
                        continue;
                    }
                    CheckBox chk = (CheckBox)item;
                    chk.IsChecked = true;
                }
            }
        }

        private void UnSelectAll(object sender, RoutedEventArgs e)
        {
            foreach (var item in grid.Children)
            {
                if (item is CheckBox)
                {
                    if (((CheckBox)item).Content.ToString() == "全选")
                    {
                        continue;
                    }
                    CheckBox chk = (CheckBox)item;
                    chk.IsChecked = false;
                }
            }
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            this.Close();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            List<string> exts = new List<string>();
            foreach (var item in grid.Children)
            {
                if (item is CheckBox)
                {
                    if (((CheckBox)item).Content.ToString() == "全选")
                    {
                        continue;
                    }
                    CheckBox chk = (CheckBox)item;
                    if (chk.IsChecked.HasValue && chk.IsChecked.Value)
                    {
                        exts.Add("*." + chk.Content.ToString());
                    }
                }
            }
            Exts = string.Join(";", exts);
            DialogResult = true;
            this.Close();
        }
    }
}
