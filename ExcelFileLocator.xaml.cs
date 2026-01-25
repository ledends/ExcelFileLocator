using FileSelector;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
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

        // 未匹配文件集合
        private HashSet<string> _unmatchedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

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
                Interval = TimeSpan.FromMilliseconds(500)
            };
            _monitorTimer.Tick += MonitorTimer_Tick;
        }

        #region Event

        // 标题栏拖动
        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        // 最小化
        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        // 最大化/还原
        private void MaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.WindowState == WindowState.Maximized)
            {
                this.WindowState = WindowState.Normal;
                btnMaximize.Content = "☐";
            }
            else
            {
                this.WindowState = WindowState.Maximized;
                btnMaximize.Content = "❐";
            }
        }

        // 关闭
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

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

        private void StartMonitoring_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_targetFolder))
            {
                ShowTopmostMessageBox("请先选择目标文件夹！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!Directory.Exists(_targetFolder))
            {
                ShowTopmostMessageBox("选择的文件夹不存在！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                _excelApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");

                if (_excelApp == null)
                {
                    ShowTopmostMessageBox("未找到运行中的Excel程序！请先打开Excel文件。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
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
                ShowTopmostMessageBox("未找到运行中的Excel程序！请先打开Excel文件。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                ShowTopmostMessageBox($"无法连接到Excel：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

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

                Excel.Range selection = _excelApp.Selection as Excel.Range;

                if (selection == null)
                {
                    return;
                }

                if (selection.Cells.Count > 1)
                {
                    UpdateExcelStatus();
                    return;
                }

                string currentAddress = selection.Address;
                string currentContent = selection.Value?.ToString() ?? "";

                UpdateExcelStatus();

                if (currentAddress != _lastCellAddress || currentContent != _lastCellContent)
                {
                    _lastCellAddress = currentAddress;
                    _lastCellContent = currentContent;

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
                AddLog("Excel已关闭，停止监控");
                StopMonitoring_Click(null, null);
            }
            catch (Exception ex)
            {
                AddLog($"监控出错: {ex.Message}");
            }
        }

        // 复制未匹配文件记录
        private void CopyUnmatched_Click(object sender, RoutedEventArgs e)
        {
            if (_unmatchedFiles.Count == 0)
            {
                ShowTopmostMessageBox("未匹配文件记录为空！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                string content = string.Join(Environment.NewLine, _unmatchedFiles.OrderBy(f => f));
                System.Windows.Clipboard.SetText(content);
                ShowTopmostMessageBox($"已复制 {_unmatchedFiles.Count} 个未匹配文件名到剪贴板", "成功",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                AddLog($"已复制 {_unmatchedFiles.Count} 个未匹配文件名");
            }
            catch (Exception ex)
            {
                ShowTopmostMessageBox($"复制失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                AddLog($"复制失败: {ex.Message}");
            }
        }

        // 清空全部（日志和未匹配记录）
        private void ClearAll_Click(object sender, RoutedEventArgs e)
        {
            // 显示确认对话框，默认按钮为"否"
            MessageBoxResult result = ShowTopmostMessageBox(
                "是否清空日志和未匹配文件记录？",
                "确认清空",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question,
                MessageBoxResult.No);  // 默认选中"否"

            if (result == MessageBoxResult.Yes)
            {
                txtLog.Clear();
                _unmatchedFiles.Clear();
                txtUnmatchedFiles.Clear();
                AddLog("日志和未匹配文件记录已清空");
            }
        }

        #endregion

        #region Private

        // 显示置顶的MessageBox（只有弹窗置顶，主窗口不置顶）
        private MessageBoxResult ShowTopmostMessageBox(string message, string title,
            MessageBoxButton button = MessageBoxButton.OK,
            MessageBoxImage icon = MessageBoxImage.Information,
            MessageBoxResult defaultResult = MessageBoxResult.None)
        {
            // 创建一个临时的隐藏置顶窗口作为MessageBox的owner
            Window dummyWindow = new Window
            {
                WindowStyle = WindowStyle.None,
                ShowInTaskbar = false,
                Width = 0,
                Height = 0,
                Left = -10000,
                Top = -10000,
                Topmost = true
            };

            dummyWindow.Show();
            dummyWindow.Activate();

            MessageBoxResult result;
            if (defaultResult != MessageBoxResult.None)
            {
                // 指定默认按钮
                result = CustomMessageBox.Show(dummyWindow, message, title, button, icon, defaultResult);
            }
            else
            {
                result = CustomMessageBox.Show(dummyWindow, message, title, button, icon);
            }

            dummyWindow.Close();

            return result;
        }

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

        private void SearchAndSelectFile(string searchPattern)
        {
            try
            {
                txtCellContent.Text = $"单元格内容: {searchPattern}";
                txtSearchPattern.Text = $"查找文件名: {searchPattern}";

                // 根据 Checkbox 状态决定搜索选项
                SearchOption searchOption = chkIncludeSubfolders.IsChecked == true
                    ? SearchOption.AllDirectories
                    : SearchOption.TopDirectoryOnly;

                var matchedFiles = Directory.GetFiles(_targetFolder, "*.*", searchOption)
                    .Where(f => Path.GetFileNameWithoutExtension(f)
                        .Equals(searchPattern, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                lstMatchedFiles.Items.Clear();

                if (matchedFiles.Any())
                {
                    txtMatchResult.Text = $"匹配结果: 找到 {matchedFiles.Count} 个文件";

                    foreach (var file in matchedFiles)
                    {
                        // 显示相对路径，更清晰
                        string relativePath = GetRelativePath(_targetFolder, file);
                        lstMatchedFiles.Items.Add($"  • {relativePath}");
                    }

                    // 打开所有匹配文件所在的文件夹并选中文件
                    SelectMultipleFilesInExplorer(matchedFiles);

                    AddLog($"✓ 找到 {matchedFiles.Count} 个文件并定位");
                }
                else
                {
                    txtMatchResult.Text = "匹配结果: 未找到匹配文件";

                    // 添加到未匹配文件记录
                    AddUnmatchedFile(searchPattern);

                    ShowTopmostMessageBox($"未找到文件名为 '{searchPattern}' 的文件", "提示",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    AddLog($"✗ 未找到文件名为 '{searchPattern}' 的文件");
                }
            }
            catch (Exception ex)
            {
                AddLog($"搜索文件时出错: {ex.Message}");
            }
        }

        // 获取相对路径（用于显示）
        private string GetRelativePath(string fromPath, string toPath)
        {
            try
            {
                Uri fromUri = new Uri(fromPath.TrimEnd('\\') + "\\");
                Uri toUri = new Uri(toPath);
                Uri relativeUri = fromUri.MakeRelativeUri(toUri);
                return Uri.UnescapeDataString(relativeUri.ToString()).Replace('/', '\\');
            }
            catch
            {
                return Path.GetFileName(toPath);
            }
        }

        // 选中多个文件（在不同文件夹中）
        private void SelectMultipleFilesInExplorer(List<string> filePaths)
        {
            // 按文件夹分组
            var filesByFolder = filePaths.GroupBy(f => Path.GetDirectoryName(f));

            foreach (var folderGroup in filesByFolder)
            {
                string folderPath = folderGroup.Key;
                var filesInFolder = folderGroup.ToList();

                // 检查该文件夹是否已打开
                var explorerWindow = FindExplorerWindow(folderPath);

                if (explorerWindow != null)
                {
                    // 窗口已打开，激活并选中第一个文件
                    AddLog($"激活已打开的文件夹: {Path.GetFileName(folderPath)}");
                    ActivateExplorerAndSelectFile(explorerWindow, filesInFolder.First());
                }
                else
                {
                    // 打开新窗口并选中第一个文件
                    AddLog($"打开文件夹: {Path.GetFileName(folderPath)}");
                    OpenNewExplorerAndSelectFile(filesInFolder.First());
                }

                // 如果同一文件夹有多个匹配文件，只打开一次窗口并选中第一个
                if (filesInFolder.Count > 1)
                {
                    AddLog($"  该文件夹包含 {filesInFolder.Count} 个匹配文件");
                }

                // 避免同时打开太多窗口，添加短暂延迟
                if (filesByFolder.Count() > 1)
                {
                    System.Threading.Thread.Sleep(300);
                }
            }
        }

        private void SelectFileInExplorer(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    ShowTopmostMessageBox("文件不存在！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string folderPath = Path.GetDirectoryName(filePath);

                var explorerWindow = FindExplorerWindow(folderPath);

                if (explorerWindow != null)
                {
                    AddLog("资源管理器窗口已打开，激活并选中文件");
                    ActivateExplorerAndSelectFile(explorerWindow, filePath);
                }
                else
                {
                    AddLog("打开新的资源管理器窗口");
                    OpenNewExplorerAndSelectFile(filePath);
                }
            }
            catch (Exception ex)
            {
                AddLog($"打开资源管理器时出错: {ex.Message}");
                ShowTopmostMessageBox($"无法打开资源管理器：{ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private dynamic FindExplorerWindow(string targetPath)
        {
            try
            {
                Type shellWindowsType = Type.GetTypeFromProgID("Shell.Application");
                dynamic shell = Activator.CreateInstance(shellWindowsType);

                string normalizedTarget = Path.GetFullPath(targetPath).TrimEnd('\\').ToLower();

                foreach (dynamic window in shell.Windows())
                {
                    try
                    {
                        string locationUrl = window.LocationURL;
                        if (string.IsNullOrEmpty(locationUrl))
                            continue;

                        if (locationUrl.StartsWith("file:///"))
                        {
                            string windowPath = Uri.UnescapeDataString(
                                locationUrl.Replace("file:///", "").Replace('/', '\\')
                            ).TrimEnd('\\').ToLower();

                            if (windowPath == normalizedTarget)
                            {
                                return window;
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                Marshal.ReleaseComObject(shell);
            }
            catch (Exception ex)
            {
                AddLog($"查找资源管理器窗口时出错: {ex.Message}");
            }

            return null;
        }

        private void ActivateExplorerAndSelectFile(dynamic explorerWindow, string filePath)
        {
            try
            {
                IntPtr hwnd = new IntPtr(explorerWindow.HWND);

                // 如果窗口最小化，先恢复
                if (NativeMethods.IsIconic(hwnd))
                {
                    NativeMethods.ShowWindow(hwnd, NativeMethods.SW_RESTORE);
                }

                // 使用更强大的窗口激活方法
                ForceWindowToForeground(hwnd);

                // 等待窗口激活
                System.Threading.Thread.Sleep(200);

                SelectFileUsingShellAPI(filePath);
            }
            catch (Exception ex)
            {
                AddLog($"激活窗口时出错: {ex.Message}");
                OpenNewExplorerAndSelectFile(filePath);
            }
            finally
            {
                try
                {
                    Marshal.ReleaseComObject(explorerWindow);
                }
                catch { }
            }
        }

        private void ForceWindowToForeground(IntPtr hwnd)
        {
            // 显示窗口
            NativeMethods.ShowWindow(hwnd, NativeMethods.SW_SHOW);

            // 模拟 Alt 键来绕过限制
            NativeMethods.keybd_event(NativeMethods.VK_MENU, 0, 0, UIntPtr.Zero);
            NativeMethods.SetForegroundWindow(hwnd);
            NativeMethods.keybd_event(NativeMethods.VK_MENU, 0, NativeMethods.KEYEVENTF_KEYUP, UIntPtr.Zero);
        }

        private void OpenNewExplorerAndSelectFile(string filePath)
        {
            System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{filePath}\"");
        }

        private void SelectFileUsingShellAPI(string filePath)
        {
            try
            {
                string folderPath = Path.GetDirectoryName(filePath);

                IntPtr folderPidl = NativeMethods.ILCreateFromPath(folderPath);
                if (folderPidl == IntPtr.Zero)
                {
                    OpenNewExplorerAndSelectFile(filePath);
                    return;
                }

                IntPtr filePidl = NativeMethods.ILCreateFromPath(filePath);
                if (filePidl == IntPtr.Zero)
                {
                    NativeMethods.ILFree(folderPidl);
                    OpenNewExplorerAndSelectFile(filePath);
                    return;
                }

                try
                {
                    IntPtr[] filePidls = new IntPtr[] { filePidl };
                    NativeMethods.SHOpenFolderAndSelectItems(folderPidl, (uint)filePidls.Length, filePidls, 0);
                }
                finally
                {
                    NativeMethods.ILFree(folderPidl);
                    NativeMethods.ILFree(filePidl);
                }
            }
            catch
            {
                OpenNewExplorerAndSelectFile(filePath);
            }
        }

        // 添加未匹配文件到记录（去重）
        private void AddUnmatchedFile(string fileName)
        {
            if (_unmatchedFiles.Add(fileName))
            {
                // 成功添加（之前不存在），更新显示
                UpdateUnmatchedFilesDisplay();
            }
        }

        // 更新未匹配文件显示
        private void UpdateUnmatchedFilesDisplay()
        {
            txtUnmatchedFiles.Text = string.Join(Environment.NewLine, _unmatchedFiles.OrderBy(f => f));
        }

        private void ClearMatchInfo()
        {
            txtCellContent.Text = "单元格内容: -";
            txtSearchPattern.Text = "查找文件名: -";
            txtMatchResult.Text = "匹配结果: -";
            lstMatchedFiles.Items.Clear();
        }

        private void AddLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            txtLog.Text += $"[{timestamp}] {message}\n";

            Dispatcher.InvokeAsync(() =>
            {
                txtLog.CaretIndex = txtLog.Text.Length;
                txtLog.ScrollToEnd();
            });
        }

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

    #endregion
}