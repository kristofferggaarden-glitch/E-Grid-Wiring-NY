using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;

namespace WpfEGridApp
{
    public class ComponentMapping
    {
        public string ExcelReference { get; set; } // F1, X2:, K3, etc.
        public int GridRow { get; set; }
        public int GridColumn { get; set; }
        public bool IsBottomSide { get; set; } // true hvis * i Excel
        public bool DefaultToBottom { get; set; } // Default side for denne mappingen
        public string Description { get; set; } // Beskrivelse for brukeren
    }

    public class ComponentMappingManager
    {
        private Dictionary<string, ComponentMapping> _mappings;
        private readonly string _mappingFileName;
        private MainWindow _mainWindow;

        public ComponentMappingManager(MainWindow mainWindow, string excelFileName)
        {
            _mainWindow = mainWindow;
            _mappings = new Dictionary<string, ComponentMapping>();

            // Lag unikt filnavn basert på Excel-fil
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelFileName);
            _mappingFileName = $"{fileNameWithoutExt}_ComponentMapping.json";

            LoadMappings();
        }

        public bool HasMapping(string excelReference)
        {
            var cleanRef = excelReference.TrimEnd('*');

            // Sjekk eksakt match først
            if (_mappings.ContainsKey(cleanRef))
                return true;

            // Sjekk prefiks match for rekkeklemmer (X20:41 skal matche X20:)
            if (cleanRef.Contains(":"))
            {
                var prefix = cleanRef.Substring(0, cleanRef.IndexOf(':') + 1);
                if (_mappings.ContainsKey(prefix))
                    return true;
            }

            return false;
        }

        public ComponentMapping GetMapping(string excelReference)
        {
            var cleanRef = excelReference.TrimEnd('*');
            var hasAsterisk = excelReference.EndsWith("*");

            // Prøv eksakt match først
            if (_mappings.TryGetValue(cleanRef, out var mapping))
            {
                return new ComponentMapping
                {
                    ExcelReference = mapping.ExcelReference,
                    GridRow = mapping.GridRow,
                    GridColumn = mapping.GridColumn,
                    IsBottomSide = hasAsterisk || mapping.DefaultToBottom,
                    DefaultToBottom = mapping.DefaultToBottom,
                    Description = mapping.Description
                };
            }

            // Prøv prefiks match for rekkeklemmer
            if (cleanRef.Contains(":"))
            {
                var prefix = cleanRef.Substring(0, cleanRef.IndexOf(':') + 1);
                if (_mappings.TryGetValue(prefix, out mapping))
                {
                    return new ComponentMapping
                    {
                        ExcelReference = mapping.ExcelReference,
                        GridRow = mapping.GridRow,
                        GridColumn = mapping.GridColumn,
                        IsBottomSide = hasAsterisk || mapping.DefaultToBottom,
                        DefaultToBottom = mapping.DefaultToBottom,
                        Description = mapping.Description
                    };
                }
            }

            return null;
        }

        public void AddMapping(string excelReference, int gridRow, int gridCol, string description = "", bool defaultToBottom = false)
        {
            var cleanRef = excelReference.TrimEnd('*');
            _mappings[cleanRef] = new ComponentMapping
            {
                ExcelReference = cleanRef,
                GridRow = gridRow,
                GridColumn = gridCol,
                IsBottomSide = false, // Settes dynamisk ved oppslag
                DefaultToBottom = defaultToBottom,
                Description = description
            };
            SaveMappings();
        }

        public void RemoveMapping(string excelReference)
        {
            var cleanRef = excelReference.TrimEnd('*');
            if (_mappings.Remove(cleanRef))
            {
                SaveMappings();
            }
        }

        public List<ComponentMapping> GetAllMappings()
        {
            return _mappings.Values.ToList();
        }

        private void SaveMappings()
        {
            try
            {
                var json = JsonSerializer.Serialize(_mappings, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(_mappingFileName, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Kunne ikke lagre mappings: {ex.Message}", "Feil",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void LoadMappings()
        {
            try
            {
                if (File.Exists(_mappingFileName))
                {
                    var json = File.ReadAllText(_mappingFileName);
                    _mappings = JsonSerializer.Deserialize<Dictionary<string, ComponentMapping>>(json)
                               ?? new Dictionary<string, ComponentMapping>();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Kunne ikke laste mappings: {ex.Message}", "Feil",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
                _mappings = new Dictionary<string, ComponentMapping>();
            }
        }
    }
}