using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace FindKeyWordsInFile
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window,INotifyPropertyChanged
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += load;
            FileList = new List<FileClass>();
            DataContext = this;
        }

        public List<string> supportExt = new List<string>();
        private void load(object sender, RoutedEventArgs e)
        {
            txtDir.Text = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
            txtDir.ToolTip = txtDir.Text;
            string exts = ConfigurationManager.AppSettings["exts"].ToString();
            string[] extsArry = exts.Split(';');
            supportExt = extsArry.Where(xz => !string.IsNullOrWhiteSpace(xz)).Select(x => "*." + x).ToList();
            //txtFilter.Text = "*.txt;*.xls;*.xlsx;*.doc;*.docx;*.h;*.config";
            txtFilter.Text = string.Join(";", supportExt);
            txtFilter.ToolTip = txtFilter.Text;
            //WordHelper.ThreadExitis("WINWORD", true);
            bar.IsEnabled = false;
            try
            {
                string enableInit = ConfigurationManager.AppSettings["enableInit"].ToString();
                string initPath = ConfigurationManager.AppSettings["initPath"].ToString();
                string initExt = ConfigurationManager.AppSettings["initExt"].ToString();
                string initKey = ConfigurationManager.AppSettings["initKey"].ToString();
                bool res = false;
                bool.TryParse(enableInit, out res);
                if (res)
                {
                    txtDir.Text = initPath;
                    txtFilter.Text = initExt;
                    txtKeyWord.Text = initKey;
                }
            }
            catch (Exception)
            {

            }
        }

        private void btnSelectDir_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog openFileDialog = new System.Windows.Forms.FolderBrowserDialog();  //选择文件夹
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtDir.Text = openFileDialog.SelectedPath;
                txtDir.ToolTip = txtDir.Text;
            }
        }

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            string dir = txtDir.Text;
            string key = txtKeyWord.Text.Trim().ToLower();
            string filter = txtFilter.Text.ToLower();
            if (string.IsNullOrEmpty(dir))
            {
                MessageBox.Show("请输入目录");
                return;
            }
            if (!Directory.Exists(dir))
            {
                MessageBox.Show("目录不存在");
                return;
            }

            if (string.IsNullOrEmpty(key))
            {
                MessageBox.Show("请输入关键词");
                return;
            }
            if (string.IsNullOrEmpty(filter))
            {
                filter = "*.txt";
            }
            string[] filterArry = filter.Split(';');
            List<string> allFiles = new List<string>();
            if (filterArry.Contains("*.doc") || filterArry.Contains("*.docx") || filterArry.Contains("*.*"))
            {
                System.Diagnostics.Process[] processList = System.Diagnostics.Process.GetProcesses();
                if (processList.Count(x => x.ProcessName == "WINWORD") > 0)
                {
                    MessageBoxResult res = MessageBox.Show($"检测到有Word程序处于打开状态或者有Word进程残留，程序运行期间Microsoft Office Word不可用，请确认所有Word已保存并关闭，继续将强制关闭Word程序,是否继续？", "", MessageBoxButton.YesNo);
                    if (res == MessageBoxResult.Cancel || res == MessageBoxResult.No || res == MessageBoxResult.None)
                    {
                        return;
                    }
                    else if (res == MessageBoxResult.Yes || res == MessageBoxResult.OK)
                    {
                        WordHelper.ThreadExitis("WINWORD", true);
                    }
                }
            }
            if (filterArry.Contains("*.*"))
            {
                string[] oneExt = Directory.GetFiles(dir, "*.*", SearchOption.AllDirectories);
                allFiles.AddRange(oneExt);
            }
            else
            {
                foreach (string ext in filterArry)
                {
                    string[] oneExt = Directory.GetFiles(dir, ext, SearchOption.AllDirectories);
                    allFiles.AddRange(oneExt);
                }
            }
            allFiles = allFiles.Distinct().ToList();
            if (allFiles.Count() > 1000)
            {
                MessageBoxResult dialogRes = MessageBox.Show($"文件数量为{allFiles.Count()},可能耗费较长时间，是否继续 ？", "", MessageBoxButton.YesNo);
                if (dialogRes == MessageBoxResult.Cancel || dialogRes == MessageBoxResult.No || dialogRes == MessageBoxResult.None)
                {
                    return;
                }
            }

            EnableControl(false);
            taskTokenSource = new CancellationTokenSource();
            bar.Maximum = allFiles.Count();
            double max = bar.Maximum;
            double curValue = 0;
            SetBarValue(0, max);
            Task task = new Task(() =>
            {
                try
                {
                    List<FileClass> UpdateList = new List<FileClass>();
                    SynDataGrid(UpdateList);
                    foreach (string item in allFiles)
                    {
                        string itemExt = System.IO.Path.GetExtension(item).ToLower();
                        if (!supportExt.Contains("*" + itemExt))
                        {
                            SetBarValue(++curValue, max);
                            continue;
                        }
                        if (taskTokenSource.Token.IsCancellationRequested)
                        {
                            SynDataGrid(UpdateList);
                            EnableControl(true);
                            return;
                        }
                        string contentPre = "";
                        string keyWord = "";
                        string contentSuf = "";
                        bool containContent = ReadFile.Read(item, key, out contentPre, out keyWord, out contentSuf);
                        if (containContent)
                        {
                            UpdateList.Add(new FileClass(item, contentPre, keyWord, contentSuf));
                        }
                        SynDataGrid(UpdateList);
                        SetBarValue(++curValue, max);
                    }
                    EnableControl(true);
                    if (UpdateList.Count() == 0)
                    {
                        MessageBox.Show("未找到关键词");
                        return;
                    }
                    //DataGridRefreshEvent();
                    SynDataGrid(UpdateList);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }, taskTokenSource.Token);
            task.Start();
        }

        private void SynDataGrid(List<FileClass> updateList)
        {
            FileList = updateList.ToList();
        }


        public delegate void FileListClear();
        public FileListClear FileListClearEvent;
        public delegate void FileListAdd(FileClass file);
        public FileListAdd FileListAddEvent;
        public delegate void DateGridRefresh();
        public DateGridRefresh DataGridRefreshEvent;
        private void SetBarValue(double v, double max)
        {
            bar.Dispatcher.Invoke(new Action<System.Windows.DependencyProperty, object>(bar.SetValue), ProgressBar.ValueProperty, v);
            txtPer.Dispatcher.Invoke(new Action(() => { txtPer.Text = (Math.Round(v / max, 4) * 100).ToString() + "%"; }));
        }

        private void EnableControl(bool value)
        {
            txtDir.Dispatcher.BeginInvoke(new Action(() =>
            {
                txtDir.IsEnabled = value;
            }));
            btnSelectDir.Dispatcher.BeginInvoke(new Action(() =>
            {
                btnSelectDir.IsEnabled = value;
            }));
            btnSelectExt.Dispatcher.BeginInvoke(new Action(() =>
            {
                btnSelectExt.IsEnabled = value;
            }));
            txtFilter.Dispatcher.BeginInvoke(new Action(() =>
            {
                txtFilter.IsEnabled = value;
            }));
            txtKeyWord.Dispatcher.BeginInvoke(new Action(() =>
            {
                txtKeyWord.IsEnabled = value;
            }));
            btnSearch.Dispatcher.BeginInvoke(new Action(() =>
            {
                btnSearch.IsEnabled = value;
            }));
            dgResult.Dispatcher.BeginInvoke(new Action(() =>
            {
                dgResult.IsEnabled = value;
            }));
        }
        private readonly object _busABcDatasLock = new object();
        public List<FileClass> _fileList;

        public List<FileClass> FileList
        {
            get { return _fileList; }
            set
            {
                _fileList = value;
                OnPropertyChanged("FileList");
            }
        }

        CancellationTokenSource taskTokenSource;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = this.PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private void dgResult_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
        }

        private void btnStop_Click(object sender, RoutedEventArgs e)
        {
            if (taskTokenSource == null)
            {
                return;
            }
            taskTokenSource.Cancel();
            EnableControl(true);
            SetBarValue(0, 1);
        }

        private void btnSelectExt_Click(object sender, RoutedEventArgs e)
        {
            SelectExt selectExtDialog = new SelectExt(txtFilter.Text);
            selectExtDialog.Owner = this;
            selectExtDialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            bool? res = selectExtDialog.ShowDialog();
            if (res.HasValue && res.Value)
            {
                string exts = selectExtDialog.Exts;
                txtFilter.Text = exts;
                txtFilter.ToolTip = txtFilter.Text;
            }
        }

        private void btnIntroduce_Click(object sender, RoutedEventArgs e)
        {
            string dir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            System.Diagnostics.Process.Start("notepad.exe", System.IO.Path.Combine(dir, "ReadMe.txt"));
        }


        private void dbClickDown(object sender, MouseButtonEventArgs e)
        {
            switch (e.ClickCount)
            {
                case 2://双击
                    {
                        string path = "" ;
                        if (sender is TextBlock)
                        {
                            TextBlock txt = sender as TextBlock;
                            path = txt.Text;
                        }
                        else if (sender is StackPanel)
                        {
                            StackPanel stackPanel = sender as StackPanel;
                            path = stackPanel.Tag.ToString();
                        }
                        if (!string.IsNullOrEmpty(path) && File.Exists(path))
                        {
                            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Explorer.exe");
                            psi.Arguments = "/e,/select," + path;
                            System.Diagnostics.Process.Start(psi);
                        }
                        break;
                    }
            }
            e.Handled = true;
        }
    }


    public class FileClass
    {
        public FileClass(string fileName, string contentPre, string keyWord,string contentSuf)
        {
            FileName = fileName;
            ContentPre = contentPre;
            KeyWord = keyWord;
            ContentSuf = contentSuf;
        }
        public string FileName { get; set; }
        public string ContentPre { get; set; }
        public string KeyWord { get; set; }
        public string ContentSuf { get; set; }
    }
}
