using Microsoft.Win32; // für OpenFileDialog
using System;
using System.Data;
using System.Windows;
using ClosedXML.Excel;
using MySql.Data.MySqlClient;

namespace ExcelToMySQL
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        // Durchsuchen-Button Klick: Öffnet einen File Dialog, um die Excel-Datei auszuwählen
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Dateien (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                filePathTextBox.Text = openFileDialog.FileName;
            }
        }

        // Importieren-Button Klick: Startet den Import der Excel-Datei in die MySQL-Datenbank
        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            string filePath = filePathTextBox.Text;
            string connectionString = connectionStringTextBox.Text;

            if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(connectionString))
            {
                MessageBox.Show("Bitte alle Felder ausfüllen!", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // Excel-Datei in DataTable einlesen
                DataTable dataTable = ReadExcelFile(filePath);

                // Daten in die MySQL-Datenbank einfügen
                InsertDataIntoDatabase(dataTable, connectionString);

                statusLabel.Content = "Status: Import erfolgreich!";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fehler: " + ex.Message, "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Methode zum Einlesen der Excel Datei in ein DataTable
        private DataTable ReadExcelFile(string excelFilePath)
        {
            DataTable dataTable = new DataTable();

            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1); // Erste Tabelle der Excel Datei
                bool firstRow = true;

                foreach (var row in worksheet.Rows())
                {
                    if (firstRow)
                    {
                        // Spaltenüberschriften hinzufügen
                        foreach (var cell in row.Cells())
                        {
                            dataTable.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        // Neue Zeile mit den Zellenwerten erstellen
                        DataRow dataRow = dataTable.NewRow();
                        int i = 0;
                        foreach (var cell in row.Cells())
                        {
                            dataRow[i] = cell.Value.ToString();
                            i++;
                        }
                        dataTable.Rows.Add(dataRow);
                    }
                }
            }

            return dataTable;
        }

        // Methode zum Einfügen der Daten in die MySQL-Datenbank
        private void InsertDataIntoDatabase(DataTable dataTable, string connectionString)
        {
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();

                foreach (DataRow row in dataTable.Rows)
                {
                    string query = "INSERT INTO content (lyrics, chords, notes, drumnotes, regieinfo, orientation) VALUES (@Wert1, @Wert2, @Wert3, @Wert4, @Wert5, @Wert6)";

                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        // Parameter aus der Zeile festlegen
                        cmd.Parameters.AddWithValue("@Wert1", row[0]);
                        cmd.Parameters.AddWithValue("@Wert2", row[1]);
                        cmd.Parameters.AddWithValue("@Wert3", row[2]);
                        cmd.Parameters.AddWithValue("@Wert4", row[3]);
                        cmd.Parameters.AddWithValue("@Wert5", row[4]);
                        cmd.Parameters.AddWithValue("@Wert6", row[5]);
                        // Weitere Parameter hinzufügen...

                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}
