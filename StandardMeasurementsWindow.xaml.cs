using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfEGridApp
{
    public partial class StandardMeasurementsWindow : Window
    {
        public ObservableCollection<StandardMeasurement> Measurements { get; set; }
        private MainWindow _parentWindow;
        private const string SettingsFileName = "StandardMeasurements.xml";

        public StandardMeasurementsWindow(MainWindow parent)
        {
            InitializeComponent();
            _parentWindow = parent;
            Measurements = new ObservableCollection<StandardMeasurement>();
            MeasurementsList.ItemsSource = Measurements;
            LoadSettingsFromFile();
            UpdateStatus("Standard målinger lastet inn");

            // Set focus to first input field
            NewColumnBText.Focus();
        }

        private void NewDistance_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AddAndApplyMeasurement();
                e.Handled = true;
            }
        }

        private void AddAndApplyMeasurement()
        {
            if (string.IsNullOrWhiteSpace(NewColumnBText.Text) ||
                string.IsNullOrWhiteSpace(NewColumnCText.Text) ||
                string.IsNullOrWhiteSpace(NewDistance.Text))
            {
                UpdateStatus("Fyll ut alle felt for å legge til en ny måling");
                return;
            }

            if (!double.TryParse(NewDistance.Text, out double distance))
            {
                UpdateStatus("Lengde må være et gyldig tall");
                return;
            }

            var measurement = new StandardMeasurement
            {
                ColumnBText = NewColumnBText.Text.Trim(),
                ColumnCText = NewColumnCText.Text.Trim(),
                Distance = distance,
                IsEnabled = true
            };

            Measurements.Add(measurement);

            // Auto-save settings
            SaveSettingsToFile();

            UpdateStatus($"La til måling: {measurement.ColumnBText} + {measurement.ColumnCText} = {measurement.Distance}mm");

            // Clear input fields and focus back to first field
            NewColumnBText.Clear();
            NewColumnCText.Clear();
            NewDistance.Clear();
            NewColumnBText.Focus();
        }

        private void DeleteMeasurement_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement button && button.Tag is StandardMeasurement measurement)
            {
                Measurements.Remove(measurement);
                SaveSettingsToFile(); // Auto-save after deletion
                UpdateStatus($"Slettet måling: {measurement.ColumnBText} + {measurement.ColumnCText}");
            }
        }

        private void SaveSettingsToFile()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(StandardMeasurement[]));
                using (var writer = new StreamWriter(SettingsFileName))
                {
                    serializer.Serialize(writer, Measurements.ToArray());
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved lagring: {ex.Message}");
            }
        }

        private void LoadSettingsFromFile()
        {
            try
            {
                if (!File.Exists(SettingsFileName))
                {
                    // Add some default measurements as examples
                    Measurements.Clear();
                    Measurements.Add(new StandardMeasurement
                    {
                        ColumnBText = "A1:1",
                        ColumnCText = "X1:",
                        Distance = 1000,
                        IsEnabled = true
                    });
                    UpdateStatus("Opprettet standard eksempel-målinger");
                    return;
                }

                var serializer = new XmlSerializer(typeof(StandardMeasurement[]));
                using (var reader = new StreamReader(SettingsFileName))
                {
                    var loadedMeasurements = (StandardMeasurement[])serializer.Deserialize(reader);
                    Measurements.Clear();
                    foreach (var measurement in loadedMeasurements)
                    {
                        Measurements.Add(measurement);
                    }
                }
                UpdateStatus($"Lastet {Measurements.Count} målinger fra {SettingsFileName}");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved lasting: {ex.Message}");
                MessageBox.Show($"Kunne ikke laste innstillinger: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ApplyAllMeasurements_Click(object sender, RoutedEventArgs e)
        {
            if (_parentWindow == null)
            {
                UpdateStatus("Feil: Ingen tilkobling til hovedvindu");
                return;
            }

            // Check if Excel is available in parent window
            if (string.IsNullOrEmpty(_parentWindow.SelectedExcelFile))
            {
                UpdateStatus("Feil: Ingen Excel-fil er valgt i hovedvinduet");
                MessageBox.Show("Du må først velge en Excel-fil i hovedvinduet", "Ingen Excel-fil",
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var appliedCount = ApplyStandardMeasurements();
                UpdateStatus($"Ferdig! Anvendt {appliedCount} målinger");
                MessageBox.Show($"Fant og la inn {appliedCount} standard målinger i Excel-filen", "Ferdig",
                               MessageBoxButton.OK, MessageBoxImage.Information);

                // Close window after applying all measurements
                this.Close();
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved anvendelse: {ex.Message}");
                MessageBox.Show($"Feil under anvendelse av målinger: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private int ApplySpecificMeasurement(StandardMeasurement measurement)
        {
            var appliedCount = 0;

            if (!measurement.IsEnabled)
            {
                UpdateStatus("Målingen er ikke aktiv");
                return 0;
            }

            // Access Excel through reflection since we don't have direct access to the worksheet
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Try to get existing Excel connection from parent window
                var excelAppField = typeof(MainWindow).GetField("excelApp",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                var worksheetField = typeof(MainWindow).GetField("worksheet",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                if (excelAppField?.GetValue(_parentWindow) is Excel.Application parentExcelApp &&
                    worksheetField?.GetValue(_parentWindow) is Excel.Worksheet parentWorksheet)
                {
                    excelApp = parentExcelApp;
                    worksheet = parentWorksheet;
                }
                else
                {
                    throw new InvalidOperationException("Kunne ikke få tilgang til Excel fra hovedvinduet");
                }

                // Find the used range to determine how far to scan
                var usedRange = worksheet.UsedRange;
                var lastRow = usedRange?.Rows?.Count ?? 100; // Default to 100 if can't determine

                // Scan through the Excel sheet
                for (int row = 2; row <= lastRow; row++) // Start from row 2
                {
                    try
                    {
                        var cellB = (worksheet.Cells[row, 2] as Excel.Range)?.Value?.ToString() ?? "";
                        var cellC = (worksheet.Cells[row, 3] as Excel.Range)?.Value?.ToString() ?? "";
                        var cellA = (worksheet.Cells[row, 1] as Excel.Range)?.Value;

                        // Skip if column A already has a value
                        if (cellA != null && !string.IsNullOrEmpty(cellA.ToString()))
                            continue;

                        // Check this specific measurement for matches
                        bool columnBMatches = string.IsNullOrEmpty(measurement.ColumnBText) ||
                                            cellB.Contains(measurement.ColumnBText, StringComparison.OrdinalIgnoreCase);
                        bool columnCMatches = string.IsNullOrEmpty(measurement.ColumnCText) ||
                                            cellC.Contains(measurement.ColumnCText, StringComparison.OrdinalIgnoreCase);

                        if (columnBMatches && columnCMatches)
                        {
                            // Apply the measurement
                            worksheet.Cells[row, 1] = measurement.Distance;
                            appliedCount++;
                            UpdateStatus($"Rad {row}: La inn {measurement.Distance}mm (matchet: '{cellB}' og '{cellC}')");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Continue with next row if there's an error with this row
                        UpdateStatus($"Advarsel rad {row}: {ex.Message}");
                    }
                }
            }
            finally
            {
                // Don't cleanup Excel objects here since they belong to the parent window
            }

            return appliedCount;
        }

        private int ApplyStandardMeasurements()
        {
            var appliedCount = 0;
            var enabledMeasurements = Measurements.Where(m => m.IsEnabled).ToList();

            if (!enabledMeasurements.Any())
            {
                UpdateStatus("Ingen aktive målinger å anvende");
                return 0;
            }

            // Access Excel through reflection since we don't have direct access to the worksheet
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Try to get existing Excel connection from parent window
                var excelAppField = typeof(MainWindow).GetField("excelApp",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                var worksheetField = typeof(MainWindow).GetField("worksheet",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                if (excelAppField?.GetValue(_parentWindow) is Excel.Application parentExcelApp &&
                    worksheetField?.GetValue(_parentWindow) is Excel.Worksheet parentWorksheet)
                {
                    excelApp = parentExcelApp;
                    worksheet = parentWorksheet;
                }
                else
                {
                    throw new InvalidOperationException("Kunne ikke få tilgang til Excel fra hovedvinduet");
                }

                // Find the used range to determine how far to scan
                var usedRange = worksheet.UsedRange;
                var lastRow = usedRange?.Rows?.Count ?? 100; // Default to 100 if can't determine

                // Scan through the Excel sheet
                for (int row = 2; row <= lastRow; row++) // Start from row 2
                {
                    try
                    {
                        var cellB = (worksheet.Cells[row, 2] as Excel.Range)?.Value?.ToString() ?? "";
                        var cellC = (worksheet.Cells[row, 3] as Excel.Range)?.Value?.ToString() ?? "";
                        var cellA = (worksheet.Cells[row, 1] as Excel.Range)?.Value;

                        // Skip if column A already has a value
                        if (cellA != null && !string.IsNullOrEmpty(cellA.ToString()))
                            continue;

                        // Check each enabled measurement for matches
                        foreach (var measurement in enabledMeasurements)
                        {
                            bool columnBMatches = string.IsNullOrEmpty(measurement.ColumnBText) ||
                                                cellB.Contains(measurement.ColumnBText, StringComparison.OrdinalIgnoreCase);
                            bool columnCMatches = string.IsNullOrEmpty(measurement.ColumnCText) ||
                                                cellC.Contains(measurement.ColumnCText, StringComparison.OrdinalIgnoreCase);

                            if (columnBMatches && columnCMatches)
                            {
                                // Apply the measurement
                                worksheet.Cells[row, 1] = measurement.Distance;
                                appliedCount++;
                                UpdateStatus($"Rad {row}: La inn {measurement.Distance}mm (matchet: '{cellB}' og '{cellC}')");
                                break; // Only apply first match per row
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Continue with next row if there's an error with this row
                        UpdateStatus($"Advarsel rad {row}: {ex.Message}");
                    }
                }
            }
            finally
            {
                // Don't cleanup Excel objects here since they belong to the parent window
            }

            return appliedCount;
        }

        private void UpdateStatus(string message)
        {
            StatusText.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            // Save settings when closing
            try
            {
                SaveSettingsToFile();
            }
            catch
            {
                // Ignore save errors on close
            }
            base.OnClosing(e);
        }
    }

    [Serializable]
    public class StandardMeasurement : INotifyPropertyChanged
    {
        private string _columnBText = "";
        private string _columnCText = "";
        private double _distance = 0;
        private bool _isEnabled = true;

        public string ColumnBText
        {
            get => _columnBText;
            set
            {
                _columnBText = value;
                OnPropertyChanged(nameof(ColumnBText));
            }
        }

        public string ColumnCText
        {
            get => _columnCText;
            set
            {
                _columnCText = value;
                OnPropertyChanged(nameof(ColumnCText));
            }
        }

        public double Distance
        {
            get => _distance;
            set
            {
                _distance = value;
                OnPropertyChanged(nameof(Distance));
            }
        }

        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                _isEnabled = value;
                OnPropertyChanged(nameof(IsEnabled));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}