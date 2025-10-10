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

                //if (success)
                //{
                //    // Lire le JSON de sortie
                //    ProcessPythonOutput();
                //}
                //else
                //{
                //    MessageBox.Show("Erreur lors de l'exécution du script Python execute",
                //        "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                //}
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
                // ✅ Chemin FIXE vers le dossier Addins de votre plugin
                string assemblyDir = @"C:\ProgramData\Autodesk\Revit\Addins\2026\BimToMode";

                Debug.WriteLine($"📁 Dossier plugin : {assemblyDir}");

                string exePath = Path.Combine(assemblyDir, "PythonScripts", "ScanToBIM.exe");

                Debug.WriteLine($"🔍 Recherche de : {exePath}");

                if (!File.Exists(exePath))
                {
                    // Message plus détaillé
                    MessageBox.Show(
                        $"L'exécutable Python est introuvable.\n\n" +
                        $"Chemin attendu :\n{exePath}\n\n" +
                        $"Vérifiez que le post-build event a bien copié le fichier.",
                        "Fichier manquant",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return false;
                }

                Debug.WriteLine($"✅ EXE trouvé, lancement...");

                ProcessStartInfo start = new ProcessStartInfo
                {
                    FileName = exePath,
                    UseShellExecute = true,
                    WorkingDirectory = Path.GetDirectoryName(exePath)
                };

                Process.Start(start);

                MessageBox.Show(
                    "L'interface Python est en cours de lancement ce processus peut prendre quelques secondes.\n\n" +
                    "Complétez le workflow, puis cliquez sur 'Suivant'.",
                    "Information",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"❌ Erreur : {ex.Message}");
                MessageBox.Show(
                    $"Erreur lors du lancement :\n{ex.Message}",
                    "Erreur",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
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