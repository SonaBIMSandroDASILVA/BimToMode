using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Newtonsoft.Json.Linq;

namespace BimToMode
{
    public partial class MainWindow : Window
    {
        private UIDocument _uiDocument;
        private Document _document;
        private string _pythonOutputPath;

        public MainWindow(UIDocument uiDocument)
        {
            InitializeComponent();
            _uiDocument = uiDocument;
            _document = _uiDocument.Document;

            _pythonOutputPath = Path.Combine(Path.GetTempPath(), "wall_detection_output.json");
        }

        private void ExecutePythonButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnExecutePython.IsEnabled = false;
                btnExecutePython.Content = "⏳ Traitement en cours...";

                // Exécuter Python
                bool success = RunPythonScript();

                if (success)
                {
                    // Lire le JSON de sortie
                    ProcessPythonOutput();
                }
                else
                {
                    MessageBox.Show("Erreur lors de l'exécution du script Python execute",
                        "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnExecutePython.IsEnabled = true;
                btnExecutePython.Content = "🔍 Lancer la Détection";
            }
        }

        private bool RunPythonScript()
        {
            try
            {
                // Récupérer le chemin de l'assembly
                string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string? assemblyDir = Path.GetDirectoryName(assemblyPath);

                if (string.IsNullOrEmpty(assemblyDir))
                {
                    Debug.WriteLine("Impossible de déterminer le dossier de l'assembly.");
                    return false;
                }

                // Chemin vers l'EXE Python
                string exePath = Path.Combine(assemblyDir, "PythonScripts", "ScanToBIM.exe");

                // Vérifier que l'EXE existe
                if (!File.Exists(exePath))
                {
                    Debug.WriteLine($"EXE introuvable : {exePath}");
                    MessageBox.Show(
                        $"L'exécutable Python est introuvable :\n{exePath}\n\n" +
                        $"Vérifiez que ScanToBIM.exe est bien copié lors du build.",
                        "Fichier manquant",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return false;
                }

                Debug.WriteLine($"Lancement de : {exePath}");

                // Lancer l'EXE Python (sans attendre)
                ProcessStartInfo start = new ProcessStartInfo
                {
                    FileName = exePath,
                    UseShellExecute = true,
                    WorkingDirectory = Path.GetDirectoryName(exePath)
                };

                Process? pythonProcess = Process.Start(start);

                if (pythonProcess == null)
                {
                    Debug.WriteLine("Impossible de démarrer le processus Python.");
                    return false;
                }

                Debug.WriteLine($"✅ Processus Python lancé (PID: {pythonProcess.Id})");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Erreur lors du lancement : {ex.Message}");
                return false;
            }
        }
        private void SuivantButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnSuivant.IsEnabled = false;
                btnSuivant.Content = "⏳ Traitement...";

                // Vérifier si le JSON existe
                if (!File.Exists(_pythonOutputPath))
                {
                    MessageBox.Show(
                        $"Le fichier de résultats n'a pas été trouvé.\n\n" +
                        $"Attendu : {_pythonOutputPath}\n\n" +
                        $"Assurez-vous d'avoir terminé le workflow dans l'interface Python " +
                        $"et cliqué sur 'Terminer'.",
                        "Fichier manquant",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);

                    btnSuivant.IsEnabled = true;
                    btnSuivant.Content = "➡️ Suivant (Importer les résultats)";
                    return;
                }

                Debug.WriteLine($"✅ JSON trouvé : {_pythonOutputPath}");

                // Lire et traiter le JSON
                ProcessPythonOutput();

                // Réinitialiser les boutons
                btnSuivant.Content = "✅ Résultats importés";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);

                btnSuivant.IsEnabled = true;
                btnSuivant.Content = "➡️ Suivant (Importer les résultats)";
            }
        }

        private void ProcessPythonOutput()
        {
            try
            {
                string jsonContent = File.ReadAllText(_pythonOutputPath);
                JObject data = JObject.Parse(jsonContent);

                MessageBox.Show($"Résultats reçus !\n\n{jsonContent}",
                    "Succès", MessageBoxButton.OK, MessageBoxImage.Information);

                // TODO: Créer les murs dans Revit avec les données

                File.Delete(_pythonOutputPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lecture JSON : {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}