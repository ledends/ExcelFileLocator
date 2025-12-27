using FileSelector;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using MessageBox = System.Windows.MessageBox;

namespace ExcelFileLocator
{
    public partial class MainWindow : Window
    {
        private Excel.Application _excelApp;
        private DispatcherTimer _monitorTimer;
        private string _targetFolder;
        private string _lastCellAddress;
        private string _lastCellContent;

        public MainWindow()
        {
            InitializeComponent();
            InitializeTimer();
            AddLog("程序已启动");
        }

        private void InitializeTimer()
        {
            _monitorTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(500) // 每500ms检查一次
            };
            _monitorTimer.Tick += MonitorTimer_Tick;
        }

        // 浏览文件夹
        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "选择要监控的文件夹",
                ShowNewFolderButton = false
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _targetFolder = dialog.SelectedPath;
                txtFolderPath.Text = _targetFolder;
                AddLog($"已选择目标文件夹: {_targetFolder}");
            }
        }

        // 开始监控
        private void StartMonitoring_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_targetFolder))
            {
                MessageBox.Show("请先选择目标文件夹！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!Directory.Exists(_targetFolder))
            {
                MessageBox.Show("选择的文件夹不存在！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // 获取正在运行的Excel实例
                _excelApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");

                if (_excelApp == null)
                {
                    MessageBox.Show("未找到运行中的Excel程序！请先打开Excel文件。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                _monitorTimer.Start();
                btnStart.IsEnabled = false;
                btnStop.IsEnabled = true;

                AddLog("开始监控Excel...");
                UpdateExcelStatus();
            }
            catch (COMException)
            {
                MessageBox.Show("未找到运行中的Excel程序！请先打开Excel文件。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法连接到Excel：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 停止监控
        private void StopMonitoring_Click(object sender, RoutedEventArgs e)
        {
            _monitorTimer.Stop();
            _excelApp = null;
            _lastCellAddress = null;
            _lastCellContent = null;

            btnStart.IsEnabled = true;
            btnStop.IsEnabled = false;

            txtExcelStatus.Text = "当前Excel: 未连接";
            txtFileName.Text = "├─ 文件名: -";
            txtSheetName.Text = "├─ 工作表: -";
            txtCurrentCell.Text = "└─ 选中单元格: -";

            ClearMatchInfo();
            AddLog("已停止监控");
        }

        // 定时器检查
        private void MonitorTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (_excelApp == null || _excelApp.Workbooks.Count == 0)
                {
                    return;
                }

                var activeWorkbook = _excelApp.ActiveWorkbook;
                var activeSheet = _excelApp.ActiveSheet as Excel.Worksheet;

                if (activeWorkbook == null || activeSheet == null)
                {
                    return;
                }

                // 获取当前选中区域
                Excel.Range selection = _excelApp.Selection as Excel.Range;

                if (selection == null)
                {
                    return;
                }

                // 检查是否为单个单元格选择（排除范围选择）
                if (selection.Cells.Count > 1)
                {
                    // 范围选择时仅更新状态，不触发查找
                    UpdateExcelStatus();
                    return;
                }

                string currentAddress = selection.Address;
                string currentContent = selection.Value?.ToString() ?? "";

                // 更新Excel状态显示
                UpdateExcelStatus();

                // 如果单元格地址或内容发生变化（单击切换）
                if (currentAddress != _lastCellAddress || currentContent != _lastCellContent)
                {
                    _lastCellAddress = currentAddress;
                    _lastCellContent = currentContent;

                    // 执行文件查找
                    if (!string.IsNullOrWhiteSpace(currentContent))
                    {
                        SearchAndSelectFile(currentContent);
                    }
                    else
                    {
                        ClearMatchInfo();
                        AddLog("当前单元格为空");
                    }
                }
            }
            catch (COMException)
            {
                // Excel已关闭，自动停止监控
                AddLog("Excel已关闭，停止监控");
                StopMonitoring_Click(null, null);
            }
            catch (Exception ex)
            {
                AddLog($"监控出错: {ex.Message}");
            }
        }

        // 更新Excel状态显示
        private void UpdateExcelStatus()
        {
            try
            {
                if (_excelApp != null && _excelApp.Workbooks.Count > 0)
                {
                    var activeWorkbook = _excelApp.ActiveWorkbook;
                    var activeSheet = _excelApp.ActiveSheet as Excel.Worksheet;
                    Excel.Range selection = _excelApp.Selection as Excel.Range;

                    txtExcelStatus.Text = "当前Excel: 已连接";
                    txtFileName.Text = $"├─ 文件名: {activeWorkbook?.Name ?? "-"}";
                    txtSheetName.Text = $"├─ 工作表: {activeSheet?.Name ?? "-"}";

                    if (selection != null && selection.Cells.Count == 1)
                    {
                        txtCurrentCell.Text = $"└─ 选中单元格: {selection.Address}";
                    }
                    else if (selection != null && selection.Cells.Count > 1)
                    {
                        txtCurrentCell.Text = $"└─ 选中单元格: {selection.Address} (范围选择，不触发查找)";
                    }
                    else
                    {
                        txtCurrentCell.Text = "└─ 选中单元格: -";
                    }
                }
            }
            catch { }
        }

        // 查找并选中文件
        private void SearchAndSelectFile(string searchPattern)
        {
            try
            {
                txtCellContent.Text = $"单元格内容: {searchPattern}";
                txtSearchPattern.Text = $"查找文件名: *{searchPattern}*";

                // 搜索匹配的文件（文件名包含单元格内容）
                var matchedFiles = Directory.GetFiles(_targetFolder, "*.*", SearchOption.TopDirectoryOnly)
                    .Where(f => Path.GetFileNameWithoutExtension(f)
                        .Contains(searchPattern, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                lstMatchedFiles.Items.Clear();

                if (matchedFiles.Any())
                {
                    txtMatchResult.Text = $"匹配结果: 找到 {matchedFiles.Count} 个文件";

                    foreach (var file in matchedFiles)
                    {
                        lstMatchedFiles.Items.Add($"  • {Path.GetFileName(file)}");
                    }

                    // 在资源管理器中选中第一个匹配的文件
                    string firstFile = matchedFiles.First();
                    SelectFileInExplorer(firstFile);

                    AddLog($"✓ 找到并定位文件: {Path.GetFileName(firstFile)}");
                }
                else
                {
                    txtMatchResult.Text = "匹配结果: 未找到匹配文件";
                    MessageBox.Show($"未找到包含 '{searchPattern}' 的文件", "提示",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    AddLog($"✗ 未找到匹配 '{searchPattern}' 的文件");
                }
            }
            catch (Exception ex)
            {
                AddLog($"搜索文件时出错: {ex.Message}");
            }
        }

        // 在资源管理器中选中文件
        private void SelectFileInExplorer(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("文件不存在！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 使用 /select 参数：
                // - 如果该文件夹未在资源管理器中打开，会打开新窗口并选中文件
                // - 如果该文件夹已在资源管理器中打开，会激活该窗口并选中文件
                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{filePath}\"");
            }
            catch (Exception ex)
            {
                AddLog($"打开资源管理器时出错: {ex.Message}");
                MessageBox.Show($"无法打开资源管理器：{ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 清除匹配信息
        private void ClearMatchInfo()
        {
            txtCellContent.Text = "单元格内容: -";
            txtSearchPattern.Text = "查找文件名: -";
            txtMatchResult.Text = "匹配结果: -";
            lstMatchedFiles.Items.Clear();
        }

        // 添加日志
        private void AddLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            txtLog.Text += $"[{timestamp}] {message}\n";

            // 自动滚动到底部
            Dispatcher.InvokeAsync(() =>
            {
                txtLog.CaretIndex = txtLog.Text.Length;
                txtLog.ScrollToEnd();
            });
        }

        // 窗口关闭时清理资源
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            _monitorTimer?.Stop();

            if (_excelApp != null)
            {
                try
                {
                    Marshal.ReleaseComObject(_excelApp);
                }
                catch { }
                _excelApp = null;
            }
        }
    }
}