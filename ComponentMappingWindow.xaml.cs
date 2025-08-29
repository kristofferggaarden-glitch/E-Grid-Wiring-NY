using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfEGridApp
{
    public partial class ComponentMappingWindow : Window
    {
        private MainWindow _mainWindow;
        private ComponentMappingManager _mappingManager;
        public ObservableCollection<MappingDisplayItem> MappingDisplayItems { get; set; }
        public ObservableCollection<string> UnmappedReferences { get; set; }

        public ComponentMappingWindow(MainWindow mainWindow, ComponentMappingManager mappingManager)
        {
            InitializeComponent();
            _mainWindow = mainWindow;
            _mappingManager = mappingManager;

            MappingDisplayItems = new ObservableCollection<MappingDisplayItem>();
            UnmappedReferences = new ObservableCollection<string>();

            MappingsList.ItemsSource = MappingDisplayItems;
            UnmappedReferencesList.ItemsSource = UnmappedReferences;

            LoadExistingMappings();
            UpdateStatus("Component mapping vindu åpnet");
        }

        private void LoadExistingMappings()
        {
            MappingDisplayItems.Clear();
            var mappings = _mappingManager.GetAllMappings();

            foreach (var mapping in mappings)
            {
                MappingDisplayItems.Add(new MappingDisplayItem
                {
                    ExcelReference = mapping.ExcelReference,
                    GridPosition = $"({mapping.GridRow},{mapping.GridColumn})",
                    Description = mapping.Description ?? ""
                });
            }

            UpdateStatus($"Lastet {mappings.Count} eksisterende mappings");
        }

        private void StartInteractiveMapping_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(NewExcelReference.Text))
            {
                UpdateStatus("Skriv inn Excel-referanse først");
                return;
            }

            var reference = NewExcelReference.Text.Trim();
            var description = NewDescription.Text.Trim();

            // Start interaktiv mapping i hovedvinduet
            _mainWindow.StartInteractiveMapping(reference, description, OnMappingCompleted);

            UpdateStatus($"Interaktiv mapping startet for {reference}. Klikk på grid-posisjon i hovedvinduet.");

            // Minimer dette vinduet for å gi plass til hovedvinduet
            this.WindowState = WindowState.Minimized;
        }

        private void OnMappingCompleted(string reference, string description)
        {
            // Oppdater visningen når mapping er fullført
            LoadExistingMappings();

            // Tøm input-feltene
            NewExcelReference.Clear();
            NewDescription.Clear();

            // Gjenopprett vinduet
            this.WindowState = WindowState.Normal;
            this.Activate();

            UpdateStatus($"Mapping fullført for {reference}");
        }

        private void FindUnmappedReferences_Click(object sender, RoutedEventArgs e)
        {
            UnmappedReferences.Clear();

            try
            {
                if (_mainWindow.worksheet == null)
                {
                    UpdateStatus("Ingen Excel-fil åpen");
                    return;
                }

                var foundReferences = new HashSet<string>();
                var usedRange = _mainWindow.worksheet.UsedRange;
                var lastRow = usedRange?.Rows?.Count ?? 100;

                // Skann kolonne B og C for referanser
                for (int row = 2; row <= lastRow; row++)
                {
                    var cellB = (_mainWindow.worksheet.Cells[row, 2] as Excel.Range)?.Value?.ToString() ?? "";
                    var cellC = (_mainWindow.worksheet.Cells[row, 3] as Excel.Range)?.Value?.ToString() ?? "";

                    ExtractReferences(cellB, foundReferences);
                    ExtractReferences(cellC, foundReferences);
                }

                // Filtrer bort referanser som allerede er mappet
                foreach (var reference in foundReferences.OrderBy(r => r))
                {
                    if (!_mappingManager.HasMapping(reference))
                    {
                        UnmappedReferences.Add(reference);
                    }
                }

                UpdateStatus($"Fant {UnmappedReferences.Count} umappede referanser av totalt {foundReferences.Count}");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved søking: {ex.Message}");
                MessageBox.Show($"Feil ved søking etter referanser: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExtractReferences(string cellValue, HashSet<string> references)
        {
            if (string.IsNullOrWhiteSpace(cellValue)) return;

            // Pattern for å finne referanser som F1, X2:41, K3, A1:1, etc.
            var patterns = new[]
            {
                @"[A-Z]\d+:\d+",     // X2:41, A1:1 format
                @"[A-Z]\d+",         // F1, K3 format
            };

            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(cellValue, pattern);
                foreach (Match match in matches)
                {
                    var reference = match.Value.TrimEnd('*'); // Fjern * hvis det finnes
                    references.Add(reference);
                }
            }
        }

        private void UnmappedReference_Selected(object sender, SelectionChangedEventArgs e)
        {
            if (UnmappedReferencesList.SelectedItem is string selectedReference)
            {
                NewExcelReference.Text = selectedReference;
                NewDescription.Text = GuessDescription(selectedReference);
            }
        }

        private string GuessDescription(string reference)
        {
            // Gi forslag til beskrivelse basert på referanse-pattern
            if (reference.StartsWith("F"))
                return "Sikring";
            else if (reference.StartsWith("X"))
                return "Rekkeklemme";
            else if (reference.StartsWith("K"))
                return "Kontaktor/Vern";
            else if (reference.StartsWith("A"))
                return "Overspenningsvern";
            else if (reference.StartsWith("S"))
                return "Signal";
            else
                return "";
        }

        private void DeleteMapping_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string excelReference)
            {
                var result = MessageBox.Show(
                    $"Slett mapping for {excelReference}?",
                    "Bekreft sletting",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    _mappingManager.RemoveMapping(excelReference);
                    LoadExistingMappings();
                    UpdateStatus($"Slettet mapping for {excelReference}");
                }
            }
        }

        private void ProcessAllWithMappings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var processor = new ExcelConnectionProcessor(_mainWindow, _mappingManager);
                var processedCount = processor.ProcessAllConnections();

                UpdateStatus($"Prosesserte {processedCount} ledninger");
                MessageBox.Show($"Ferdig! Prosesserte {processedCount} ledninger automatisk.",
                               "Automatisk prosessering fullført",
                               MessageBoxButton.OK, MessageBoxImage.Information);

                // Oppdater display i hovedvinduet
                _mainWindow.UpdateExcelDisplayText();
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil under prosessering: {ex.Message}");
                MessageBox.Show($"Feil under automatisk prosessering: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void UpdateStatus(string message)
        {
            StatusText.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
        }
    }

    public class MappingDisplayItem
    {
        public string ExcelReference { get; set; }
        public string GridPosition { get; set; }
        public string Description { get; set; }
    }

    // Prosessor klasse som bruker mappings til å behandle alle ledninger
    public class ExcelConnectionProcessor
    {
        private readonly MainWindow _mainWindow;
        private readonly ComponentMappingManager _mappingManager;

        public ExcelConnectionProcessor(MainWindow mainWindow, ComponentMappingManager mappingManager)
        {
            _mainWindow = mainWindow;
            _mappingManager = mappingManager;
        }

        public int ProcessAllConnections()
        {
            if (_mainWindow.worksheet == null)
                throw new InvalidOperationException("Ingen Excel-fil er åpen");

            int processedCount = 0;
            var usedRange = _mainWindow.worksheet.UsedRange;
            var lastRow = usedRange?.Rows?.Count ?? 100;

            for (int row = 2; row <= lastRow; row++)
            {
                try
                {
                    var cellB = (_mainWindow.worksheet.Cells[row, 2] as Excel.Range)?.Value?.ToString() ?? "";
                    var cellC = (_mainWindow.worksheet.Cells[row, 3] as Excel.Range)?.Value?.ToString() ?? "";
                    var cellA = (_mainWindow.worksheet.Cells[row, 1] as Excel.Range)?.Value;

                    // Skip hvis allerede har måleverdi
                    if (cellA != null && !string.IsNullOrEmpty(cellA.ToString()))
                        continue;

                    // Skip hvis ingen punkt A eller B
                    if (string.IsNullOrWhiteSpace(cellB) && string.IsNullOrWhiteSpace(cellC))
                        continue;

                    // Prosesser denne raden
                    var distance = ProcessSingleConnection(cellB, cellC, row);
                    if (distance.HasValue)
                    {
                        _mainWindow.worksheet.Cells[row, 1] = distance.Value;
                        processedCount++;
                    }
                }
                catch (Exception ex)
                {
                    // Log feil men fortsett med neste rad
                    System.Diagnostics.Debug.WriteLine($"Feil på rad {row}: {ex.Message}");
                }
            }

            return processedCount;
        }

        private double? ProcessSingleConnection(string pointAText, string pointBText, int row)
        {
            var pointA = FindConnectionPoint(pointAText);
            var pointB = FindConnectionPoint(pointBText);

            if (pointA == null || pointB == null)
                return null;

            // Finn grid-posisjoner
            var gridPosA = GetGridPosition(pointA);
            var gridPosB = GetGridPosition(pointB);

            if (!gridPosA.HasValue || !gridPosB.HasValue)
                return null;

            // Beregn avstand ved å bruke eksisterende PathFinder
            return CalculateDistance(gridPosA.Value, gridPosB.Value, pointA, pointB);
        }

        private ConnectionPoint FindConnectionPoint(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;

            // Ekstrahér første referanse fra teksten
            var patterns = new[]
            {
                @"[A-Z]\d+:\d+\*?",     // X2:41, A1:1 format (med eller uten *)
                @"[A-Z]\d+\*?",         // F1, K3 format (med eller uten *)
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern);
                if (match.Success)
                {
                    return new ConnectionPoint
                    {
                        Reference = match.Value.TrimEnd('*'),
                        IsBottomSide = match.Value.EndsWith("*"),
                        OriginalText = text
                    };
                }
            }

            return null;
        }

        private (int Row, int Col)? GetGridPosition(ConnectionPoint point)
        {
            var mapping = _mappingManager.GetMapping(point.Reference + (point.IsBottomSide ? "*" : ""));
            if (mapping == null)
            {
                // Prøv uten * også
                mapping = _mappingManager.GetMapping(point.Reference);
            }

            if (mapping != null)
            {
                return (mapping.GridRow, mapping.GridColumn);
            }

            return null;
        }

        private double CalculateDistance((int Row, int Col) posA, (int Row, int Col) posB,
                                       ConnectionPoint pointA, ConnectionPoint pointB)
        {
            // Bruk eksisterende PathFinder
            var allCells = _mainWindow.GetAllCells();
            if (!allCells.TryGetValue(posA, out var startCell) ||
                !allCells.TryGetValue(posB, out var endCell))
            {
                return 0;
            }

            var path = PathFinder.FindShortestPath(startCell, endCell, allCells,
                                                  _mainWindow.HasHorizontalNeighbor);

            if (path == null) return 0;

            // Beregn base avstand
            double baseDistance = PathFinder.CalculateDistance(path, false, _mainWindow.HasHorizontalNeighbor);

            // Legg til ekstra avstand for tilkoblinger
            double connectionDistance = GetConnectionDistance(pointA) + GetConnectionDistance(pointB);

            return baseDistance + connectionDistance;
        }

        private double GetConnectionDistance(ConnectionPoint point)
        {
            // Forskjellige komponenter har forskjellige tilkoblingslengder
            var reference = point.Reference.ToUpper();

            if (reference.StartsWith("F")) // Sikring
                return point.IsBottomSide ? 50 : 30;
            else if (reference.StartsWith("X")) // Rekkeklemme
                return point.IsBottomSide ? 40 : 20;
            else if (reference.StartsWith("K")) // Kontaktor/Vern
                return 60;
            else if (reference.StartsWith("A")) // Overspenningsvern
                return 45;
            else
                return 25; // Standard
        }
    }

    public class ConnectionPoint
    {
        public string Reference { get; set; }
        public bool IsBottomSide { get; set; }
        public string OriginalText { get; set; }
    }
}