using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfEGridApp
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private int _sections = 5;
        private int _rows = 7;
        private int _cols = 4;
        private string _selectedExcelFile;
        private string _excelDisplayText;
        private SpecialPoint _lockedPointA; // Track the locked point A (Motor or Door)
        private Button _lockedButton; // Track the lock button for text updates
        private int _currentExcelRow = 2; // Start with row 2 (B2, C2)

        // Nye fields for mapping
        private ComponentMappingManager _componentMappingManager;
        private string _currentMappingReference = "";
        private string _currentMappingDescription = "";
        private bool _isInMappingMode = false;
        private Action<string, string> _mappingCompletedCallback;

        public int Sections
        {
            get => _sections;
            set { _sections = value; OnPropertyChanged(nameof(Sections)); }
        }

        public int Rows
        {
            get => _rows;
            set { _rows = value; OnPropertyChanged(nameof(Rows)); }
        }

        public int Cols
        {
            get => _cols;
            set { _cols = value; OnPropertyChanged(nameof(Cols)); }
        }

        public string SelectedExcelFile
        {
            get => _selectedExcelFile;
            set { _selectedExcelFile = value; OnPropertyChanged(nameof(SelectedExcelFile)); }
        }

        public string ExcelDisplayText
        {
            get => _excelDisplayText;
            set { _excelDisplayText = value; OnPropertyChanged(nameof(ExcelDisplayText)); }
        }

        private readonly Dictionary<(int globalRow, int globalCol), Cell> allCells = new();
        private readonly List<SpecialPoint> doorPoints = new();
        private readonly List<SpecialPoint> motorPoints = new();
        private object startPoint;
        private object endPoint;
        private Excel.Application excelApp;
        private Excel.Workbook workbook;
        public Excel.Worksheet worksheet; // Made public for ComponentMappingManager

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            UpdateExcelDisplayText();
            BuildAllSections();
        }

        // New method to open Standard Measurements window
        private void OpenStandardMeasurements_Click(object sender, RoutedEventArgs e)
        {
            var standardMeasurementsWindow = new StandardMeasurementsWindow(this);
            standardMeasurementsWindow.Show();
        }

        // Method to open Component Mapping window
        private void OpenComponentMapping_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedExcelFile))
            {
                MessageBox.Show("Velg først en Excel-fil", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_componentMappingManager == null)
            {
                _componentMappingManager = new ComponentMappingManager(this, SelectedExcelFile);
            }

            var mappingWindow = new ComponentMappingWindow(this, _componentMappingManager);
            mappingWindow.Show();
        }

        // Method for interactive mapping
        public void StartInteractiveMapping(string excelReference, string description, Action<string, string> onCompleted)
        {
            _currentMappingReference = excelReference;
            _currentMappingDescription = description;
            _isInMappingMode = true;
            _mappingCompletedCallback = onCompleted;

            ResultText.Text = $"Klikk på grid-posisjonen for {excelReference}";

            // Change background color on all cells to show mapping mode
            foreach (var cell in allCells.Values)
            {
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(100, 120, 140));
            }

            // Give focus to main window
            this.Activate();
            this.Focus();
        }

        // Method for automatic measuring of all connections
        private void AutomaticMeasureAll_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
            {
                MessageBox.Show("Velg først en Excel-fil", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_componentMappingManager == null)
            {
                MessageBox.Show("Du må først sette opp component mappings. Bruk 'Component Mapping' knappen.",
                               "Mappings mangler", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show(
                "Dette vil automatisk måle alle ledninger som har component mappings. Fortsette?",
                "Automatisk måling",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    var processor = new ExcelConnectionProcessor(this, _componentMappingManager);
                    var processedCount = processor.ProcessAllConnections();

                    MessageBox.Show($"Automatisk måling fullført!\nProsesserte {processedCount} ledninger.",
                                   "Ferdig", MessageBoxButton.OK, MessageBoxImage.Information);

                    UpdateExcelDisplayText();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Feil under automatisk måling: {ex.Message}", "Feil",
                                   MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        public void UpdateExcelDisplayText()
        {
            if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
            {
                ExcelDisplayText = "";
                return;
            }

            // Find next available row first
            FindNextAvailableRowForDisplay();

            try
            {
                string cellB = (worksheet.Cells[_currentExcelRow, 2] as Excel.Range)?.Value?.ToString() ?? "";
                string cellC = (worksheet.Cells[_currentExcelRow, 3] as Excel.Range)?.Value?.ToString() ?? "";
                ExcelDisplayText = string.IsNullOrEmpty(cellB) && string.IsNullOrEmpty(cellC) ? "" : $"{cellB} - {cellC}".Trim();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading Excel data for row {_currentExcelRow}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ExcelDisplayText = "";
            }
        }

        // Public access to allCells for ComponentMappingManager
        public Dictionary<(int, int), Cell> GetAllCells()
        {
            return allCells;
        }

        private void FindNextAvailableRowForDisplay()
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                    return;

                // Find next row that doesn't have a measurement in column A
                while (_currentExcelRow <= 1000) // Safety limit
                {
                    var cellA = (worksheet.Cells[_currentExcelRow, 1] as Excel.Range)?.Value;
                    if (cellA == null || string.IsNullOrEmpty(cellA.ToString()))
                    {
                        // Found an empty row, check if it has content in B or C to display
                        var cellB = (worksheet.Cells[_currentExcelRow, 2] as Excel.Range)?.Value?.ToString() ?? "";
                        var cellC = (worksheet.Cells[_currentExcelRow, 3] as Excel.Range)?.Value?.ToString() ?? "";

                        if (!string.IsNullOrEmpty(cellB) || !string.IsNullOrEmpty(cellC))
                        {
                            // Found a row with content to display
                            break;
                        }
                    }
                    _currentExcelRow++;
                }
            }
            catch (Exception)
            {
                // If there's an error, just continue with current row
            }
        }

        private bool InitializeExcel(string filePath)
        {
            try
            {
                CleanupExcel();
                excelApp = new Excel.Application { Visible = true };
                if (File.Exists(filePath))
                {
                    workbook = excelApp.Workbooks.Open(filePath);
                }
                else
                {
                    workbook = excelApp.Workbooks.Add();
                }
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to initialize Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            CleanupExcel();
        }

        private void CleanupExcel()
        {
            try
            {
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error cleaning up Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                worksheet = null;
                workbook = null;
                excelApp = null;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private void Rebuild_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(SectionBox.Text, out int s) || s <= 0 || s > 10)
            {
                MessageBox.Show("Please enter a valid number of sections (1-10).", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(RowBox.Text, out int r) || r <= 0 || r > 20)
            {
                MessageBox.Show("Please enter a valid number of rows (1-20).", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(ColBox.Text, out int c) || c <= 0 || c > 10)
            {
                MessageBox.Show("Please enter a valid number of columns (1-10).", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Sections = s;
            Rows = r;
            Cols = c;
            _lockedPointA = null; // Reset lock on rebuild
            _lockedButton = null;
            _currentExcelRow = 2; // Reset to row 2
            BuildAllSections();
            UpdateExcelDisplayText();
        }

        private void SelectExcelFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                Title = "Select Excel File for Measurements"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                SelectedExcelFile = openFileDialog.FileName;
                if (InitializeExcel(SelectedExcelFile))
                {
                    MessageBox.Show($"Excel file '{SelectedExcelFile}' selected and opened.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    _currentExcelRow = 2;

                    // Initialize mapping manager for this file
                    _componentMappingManager = new ComponentMappingManager(this, SelectedExcelFile);

                    UpdateExcelDisplayText();
                }
                else
                {
                    SelectedExcelFile = null;
                    ExcelDisplayText = "";
                    _componentMappingManager = null;
                }
            }
        }

        private void DeleteLastMeasurement_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                {
                    MessageBox.Show("Please select an Excel file first.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                int lastRow = _currentExcelRow - 1; // Check the previous row
                if (lastRow >= 2 && (worksheet.Cells[lastRow, 1] as Excel.Range)?.Value != null)
                {
                    ((Excel.Range)worksheet.Cells[lastRow, 1]).Clear();
                    _currentExcelRow = lastRow; // Set to the cleared row
                    object cellValue = (worksheet.Cells[lastRow, 1] as Excel.Range)?.Value;
                    double? lastValue = cellValue != null && double.TryParse(cellValue.ToString(), out double parsedValue) ? parsedValue : null;
                    ResultText.Text = lastValue.HasValue ? $"Shortest path: {lastValue:F2} mm" : "";
                    UpdateExcelDisplayText();
                }
                else
                {
                    MessageBox.Show("No measurements to delete.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    _currentExcelRow = 2; // Reset to row 2
                    ResultText.Text = _lockedPointA != null ? $"Locked point A: {_lockedPointA.Type} {_lockedPointA.SectionIndex + 1}" : "";
                    UpdateExcelDisplayText();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error deleting last measurement: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BuildAllSections()
        {
            MainPanel.Children.Clear();
            allCells.Clear();
            doorPoints.Clear();
            motorPoints.Clear();
            ResultText.Text = "";
            startPoint = null;
            endPoint = null;

            for (int s = 0; s < Sections; s++)
            {
                var sectionPanel = new StackPanel
                {
                    Margin = new Thickness(15),
                    Orientation = Orientation.Vertical,
                    Background = Brushes.Transparent
                };
                var doorBtn = new Button
                {
                    Content = $"Door {s + 1}",
                    Width = 200,
                    Height = 40,
                    Style = (Style)FindResource("OrangeButtonStyle"),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    ToolTip = $"Select door {s + 1} as a path point",
                    Tag = s
                };
                doorBtn.Click += (sender, e) => HandlePointClick(sender, doorPoints);
                sectionPanel.Children.Add(doorBtn);
                var doorLockBtn = new Button
                {
                    Content = $"Lock Door {s + 1}",
                    Width = 200,
                    Height = 40,
                    Style = (Style)FindResource("RoundedButtonStyle"),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    ToolTip = $"Lock or unlock door {s + 1} as point A",
                    Tag = s
                };
                doorLockBtn.Click += (sender, e) => LockPointA_Click(sender, doorPoints);
                sectionPanel.Children.Add(doorLockBtn);
                var grid = new Grid
                {
                    Background = Brushes.Transparent,
                    Margin = new Thickness(0, 10, 0, 10)
                };
                sectionPanel.Children.Add(grid);
                var motorBtn = new Button
                {
                    Content = $"Motor {s + 1}",
                    Width = 200,
                    Height = 40,
                    Style = (Style)FindResource("OrangeButtonStyle"),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    ToolTip = $"Select motor {s + 1} as a path point",
                    Tag = s
                };
                motorBtn.Click += (sender, e) => HandlePointClick(sender, motorPoints);
                sectionPanel.Children.Add(motorBtn);
                var motorLockBtn = new Button
                {
                    Content = $"Lock Motor {s + 1}",
                    Width = 200,
                    Height = 40,
                    Style = (Style)FindResource("RoundedButtonStyle"),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    ToolTip = $"Lock or unlock motor {s + 1} as point A",
                    Tag = s
                };
                motorLockBtn.Click += (sender, e) => LockPointA_Click(sender, motorPoints);
                sectionPanel.Children.Add(motorLockBtn);

                grid.RowDefinitions.Clear();
                grid.ColumnDefinitions.Clear();
                for (int r = 0; r < Rows; r++)
                    grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                for (int c = 0; c < Cols; c++)
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                for (int row = 0; row < Rows; row++)
                {
                    if (row % 2 == 0) // Even-indexed rows (0, 2, 4, 6, 8, ...)
                    {
                        for (int col = 0; col < Cols; col++)
                            AddCell(grid, row, col, s);
                    }
                    else // Odd-indexed rows (1, 3, 5, 7, 9, ...)
                    {
                        AddCell(grid, row, 0, s);
                    }
                }

                doorPoints.Add(new SpecialPoint
                {
                    SectionIndex = s,
                    Type = SpecialPointType.Door,
                    Button = doorBtn,
                    GlobalRow = 0,
                    GlobalCol = s * Cols
                });
                motorPoints.Add(new SpecialPoint
                {
                    SectionIndex = s,
                    Type = SpecialPointType.Motor,
                    Button = motorBtn,
                    GlobalRow = Rows - 1,
                    GlobalCol = s * Cols + (Cols - 1)
                });
                MainPanel.Children.Add(sectionPanel);
            }

            // Restore locked point if it exists
            if (_lockedPointA != null)
            {
                var points = _lockedPointA.Type == SpecialPointType.Door ? doorPoints : motorPoints;
                var point = points.FirstOrDefault(p => p.SectionIndex == _lockedPointA.SectionIndex);
                if (point != null)
                {
                    startPoint = point;
                    point.Button.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                    ResultText.Text = $"Locked point A: {point.Type} {point.SectionIndex + 1}";
                    foreach (var panel in MainPanel.Children.OfType<StackPanel>())
                    {
                        foreach (var btn in panel.Children.OfType<Button>().Where(b => b.Tag != null && (int)b.Tag == point.SectionIndex && b.Content.ToString().StartsWith("Lock")))
                        {
                            if ((btn.Content.ToString().Contains("Door") && point.Type == SpecialPointType.Door) ||
                                (btn.Content.ToString().Contains("Motor") && point.Type == SpecialPointType.Motor))
                            {
                                _lockedButton = btn;
                                btn.Content = $"Unlock {point.Type} {point.SectionIndex + 1}";
                                break;
                            }
                        }
                    }
                }
                else
                {
                    _lockedPointA = null;
                    _lockedButton = null;
                }
            }
        }

        private void AddCell(Grid grid, int localRow, int localCol, int sectionIndex)
        {
            var btn = new Button
            {
                Width = 50,
                Height = 50,
                Margin = new Thickness(4),
                Background = new SolidColorBrush(Color.FromRgb(74, 90, 91)),
                ToolTip = $"Cell ({localRow}, {localCol}) in section {sectionIndex + 1}",
                Style = (Style)FindResource("RoundedButtonStyle")
            };
            btn.Click += Cell_Click;
            Grid.SetRow(btn, localRow);
            Grid.SetColumn(btn, localCol);
            grid.Children.Add(btn);
            int globalRow = localRow;
            int globalCol = sectionIndex * Cols + localCol;
            allCells[(globalRow, globalCol)] = new Cell(globalRow, globalCol, btn);
        }

        private void Cell_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;

            // Handle mapping mode first
            if (_isInMappingMode)
            {
                // Find grid position for this cell
                var cell = allCells.Values.FirstOrDefault(c => c.ButtonRef == btn);
                if (cell != null && !string.IsNullOrEmpty(_currentMappingReference))
                {
                    // Add mapping
                    _componentMappingManager?.AddMapping(_currentMappingReference, cell.Row, cell.Col, _currentMappingDescription);

                    MessageBox.Show($"Mappet {_currentMappingReference} til posisjon ({cell.Row}, {cell.Col})",
                                   "Mapping lagret", MessageBoxButton.OK, MessageBoxImage.Information);

                    // Call callback
                    _mappingCompletedCallback?.Invoke(_currentMappingReference, _currentMappingDescription);

                    // End mapping mode
                    _isInMappingMode = false;
                    _currentMappingReference = "";
                    _currentMappingDescription = "";
                    _mappingCompletedCallback = null;

                    // Reset grid colors
                    ResetCellColors();
                    ResultText.Text = "";
                }
                return;
            }

            // Original Cell_Click logic for normal measuring...
            if (_lockedPointA != null)
            {
                if (endPoint != null)
                    ResetSelection();
                endPoint = btn;
                btn.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
                ProcessPath();
            }
            else
            {
                if (startPoint != null && endPoint != null)
                    ResetSelection();
                if (startPoint == null)
                {
                    startPoint = btn;
                    btn.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                    ResultText.Text = "";
                }
                else if (endPoint == null)
                {
                    endPoint = btn;
                    btn.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
                    ProcessPath();
                }
            }
        }

        // Helper method to reset cell colors
        private void ResetCellColors()
        {
            foreach (var cell in allCells.Values)
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(74, 90, 91));
        }

        // Public access to HasHorizontalNeighbor
        public bool HasHorizontalNeighbor(int row, int col)
        {
            return allCells.ContainsKey((row, col - 1)) || allCells.ContainsKey((row, col + 1));
        }

        private void HandlePointClick(object sender, List<SpecialPoint> points)
        {
            var btn = sender as Button;
            if (btn.Tag == null) return;
            int sectionIndex = (int)btn.Tag;
            var special = points.FirstOrDefault(p => p.SectionIndex == sectionIndex);
            if (special == null) return;

            if (_lockedPointA != null)
            {
                if (endPoint != null)
                    ResetSelection();
                endPoint = special;
                special.Button.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
                ProcessPath();
            }
            else
            {
                if (startPoint != null && endPoint != null)
                    ResetSelection();
                if (startPoint == null)
                {
                    startPoint = special;
                    special.Button.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                    ResultText.Text = "";
                }
                else if (endPoint == null)
                {
                    endPoint = special;
                    special.Button.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
                    ProcessPath();
                }
            }
        }

        private void LockPointA_Click(object sender, List<SpecialPoint> points)
        {
            var btn = sender as Button;
            if (btn.Tag == null) return;
            int sectionIndex = (int)btn.Tag;
            var special = points.FirstOrDefault(p => p.SectionIndex == sectionIndex);
            if (special == null) return;

            if (_lockedPointA == special)
            {
                _lockedPointA = null;
                _lockedButton = null;
                btn.Content = $"Lock {special.Type} {sectionIndex + 1}";
                ResetSelection();
            }
            else
            {
                if (_lockedPointA != null && _lockedButton != null)
                {
                    _lockedButton.Content = $"Lock {_lockedPointA.Type} {_lockedPointA.SectionIndex + 1}";
                    _lockedPointA.Button.Background = (Brush)FindResource("OrangeButtonBrush");
                }
                _lockedPointA = special;
                _lockedButton = btn;
                ResetSelection();
                startPoint = special;
                special.Button.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                btn.Content = $"Unlock {special.Type} {sectionIndex + 1}";
                ResultText.Text = $"Locked point A: {special.Type} {sectionIndex + 1}";
            }
        }

        private void ProcessPath()
        {
            Cell startCell = null;
            Cell endCell = null;
            double extraDistance = 0;

            if (startPoint is SpecialPoint spStart)
            {
                startCell = allCells[(spStart.GlobalRow, spStart.GlobalCol)];
                extraDistance += spStart.Type == SpecialPointType.Door ? 1000 : 500;
            }
            else if (startPoint is Button btnStart)
            {
                startCell = allCells.Values.FirstOrDefault(c => c.ButtonRef == btnStart);
            }

            if (endPoint is SpecialPoint spEnd)
            {
                endCell = allCells[(spEnd.GlobalRow, spEnd.GlobalCol)];
                extraDistance += spEnd.Type == SpecialPointType.Door ? 1000 : 500;
            }
            else if (endPoint is Button btnEnd)
            {
                endCell = allCells.Values.FirstOrDefault(c => c.ButtonRef == btnEnd);
            }

            if (startCell == null || endCell == null)
            {
                ResultText.Text = "Invalid start or end point.";
                ResetSelection();
                return;
            }

            var path = PathFinder.FindShortestPath(startCell, endCell, allCells, HasHorizontalNeighbor);
            if (path == null)
            {
                ResultText.Text = "No valid path found.";
                ResetSelection();
                return;
            }

            bool endsInSpecial = endPoint is SpecialPoint;
            double totalDistance = PathFinder.CalculateDistance(path, endsInSpecial, HasHorizontalNeighbor) + extraDistance;
            HighlightPath(path);

            if (startPoint is SpecialPoint sp1 && endPoint is SpecialPoint sp2)
            {
                if ((sp1.Type == SpecialPointType.Motor && sp2.Type == SpecialPointType.Door) ||
                    (sp1.Type == SpecialPointType.Door && sp2.Type == SpecialPointType.Motor))
                {
                    sp1.Button.Background = new SolidColorBrush(Color.FromRgb(0, 178, 148));
                    sp2.Button.Background = new SolidColorBrush(Color.FromRgb(0, 178, 148));
                }
            }

            ResultText.Text = $"Shortest path: {totalDistance:F2} mm";

            // Skip to next available row if current row already has a measurement
            FindNextAvailableRow();

            LogMeasurementToExcel(totalDistance);
            _currentExcelRow++; // Move to next row
            UpdateExcelDisplayText();
        }

        private void FindNextAvailableRow()
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                    return;

                // Check if current row already has a measurement in column A
                while (_currentExcelRow <= 1000) // Safety limit
                {
                    var cellA = (worksheet.Cells[_currentExcelRow, 1] as Excel.Range)?.Value;
                    if (cellA == null || string.IsNullOrEmpty(cellA.ToString()))
                    {
                        // Found an empty row
                        break;
                    }
                    _currentExcelRow++; // Move to next row
                }
            }
            catch (Exception)
            {
                // If there's an error, just continue with current row
            }
        }

        private void ResetSelection()
        {
            foreach (var cell in allCells.Values)
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(74, 90, 91));
            foreach (var dp in doorPoints)
                dp.Button.Background = (Brush)FindResource("OrangeButtonBrush");
            foreach (var mp in motorPoints)
                mp.Button.Background = (Brush)FindResource("OrangeButtonBrush");
            startPoint = _lockedPointA;
            endPoint = null;
            if (_lockedPointA != null)
            {
                startPoint = _lockedPointA;
                _lockedPointA.Button.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                ResultText.Text = $"Locked point A: {_lockedPointA.Type} {_lockedPointA.SectionIndex + 1}";
            }
            else
            {
                ResultText.Text = "";
            }
        }

        private void HighlightPath(List<Cell> path)
        {
            foreach (var cell in allCells.Values)
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(74, 90, 91));
            foreach (var dp in doorPoints)
                dp.Button.Background = (Brush)FindResource("OrangeButtonBrush");
            foreach (var mp in motorPoints)
                mp.Button.Background = (Brush)FindResource("OrangeButtonBrush");
            foreach (var cell in path)
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(0, 178, 148));
            if (startPoint is Button b1)
                b1.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
            if (endPoint is Button b2)
                b2.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
            if (startPoint is SpecialPoint sp1)
                sp1.Button.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
            if (endPoint is SpecialPoint sp2)
                sp2.Button.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
        }

        private void LogMeasurementToExcel(double distance)
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                {
                    MessageBox.Show("Please select an Excel file first.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                worksheet.Cells[_currentExcelRow, 1] = distance; // Log to column A, current row
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error writing to Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    public static class PathFinder
    {
        public static List<Cell> FindShortestPath(Cell start, Cell end, Dictionary<(int, int), Cell> allCells, Func<int, int, bool> hasHorizontalNeighbor)
        {
            var dist = new Dictionary<Cell, double> { { start, 0 } };
            var prev = new Dictionary<Cell, Cell>();
            var queue = new PriorityQueue<Cell, double>();
            queue.Enqueue(start, 0);
            foreach (var cell in allCells.Values)
            {
                if (cell != start)
                    dist[cell] = double.PositiveInfinity;
                prev[cell] = null;
            }
            while (queue.Count > 0)
            {
                var u = queue.Dequeue();
                if (u == end) break;
                foreach (var neighbor in GetNeighbors(u, allCells))
                {
                    double weight = hasHorizontalNeighbor(neighbor.Row, neighbor.Col) ? 100 : 50;
                    double alt = dist[u] + weight;
                    if (alt < dist[neighbor])
                    {
                        dist[neighbor] = alt;
                        prev[neighbor] = u;
                        queue.Enqueue(neighbor, alt);
                    }
                }
            }
            if (double.IsInfinity(dist[end])) return null;
            var path = new List<Cell>();
            for (var curr = end; curr != null; curr = prev[curr])
                path.Add(curr);
            path.Reverse();
            return path;
        }

        private static List<Cell> GetNeighbors(Cell cell, Dictionary<(int, int), Cell> allCells)
        {
            var directions = new[] { (-1, 0), (1, 0), (0, -1), (0, 1) };
            var neighbors = new List<Cell>();
            foreach (var (dr, dc) in directions)
            {
                int nr = cell.Row + dr;
                int nc = cell.Col + dc;
                if (allCells.TryGetValue((nr, nc), out var neighbor))
                    neighbors.Add(neighbor);
            }
            return neighbors;
        }

        public static double CalculateDistance(List<Cell> path, bool endsInSpecial, Func<int, int, bool> hasHorizontalNeighbor)
        {
            if (path.Count == 0) return 0;
            double distance = 100; // Initial move: 10 cm = 100 mm
            for (int i = 1; i < path.Count - 1; i++)
                distance += hasHorizontalNeighbor(path[i].Row, path[i].Col) ? 100 : 50;
            if (!endsInSpecial)
                distance += 200;
            return distance;
        }
    }

    public class Cell
    {
        public int Row { get; }
        public int Col { get; }
        public Button ButtonRef { get; }
        public Cell(int row, int col, Button button)
        {
            Row = row;
            Col = col;
            ButtonRef = button;
        }
    }

    public enum SpecialPointType { Door, Motor }

    public class SpecialPoint
    {
        public int SectionIndex { get; set; }
        public SpecialPointType Type { get; set; }
        public Button Button { get; set; }
        public int GlobalRow { get; set; }
        public int GlobalCol { get; set; }
    }

    public class PriorityQueue<TItem, TPriority> where TPriority : IComparable<TPriority>
    {
        private readonly List<(TItem Item, TPriority Priority)> elements = new();
        public int Count => elements.Count;
        public void Enqueue(TItem item, TPriority priority)
        {
            elements.Add((item, priority));
        }
        public TItem Dequeue()
        {
            int bestIndex = 0;
            for (int i = 1; i < elements.Count; i++)
            {
                if (elements[i].Priority.CompareTo(elements[bestIndex].Priority) < 0)
                    bestIndex = i;
            }
            var result = elements[bestIndex].Item;
            elements.RemoveAt(bestIndex);
            return result;
        }
    }
}