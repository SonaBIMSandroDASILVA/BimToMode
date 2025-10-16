using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Media.Imaging;

namespace BimToMode
{
    public class App : IExternalApplication
    {
        private string EmbeddedParentIconLocation => GetType().Namespace + ".Ressources.Icons.Embedded.";
        string name = string.Empty;
        string text = string.Empty;
        string toolTip = string.Empty;
        string largeImageName = string.Empty;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Création d'un onglet personnalisé dans le ruban Revit
                string tabName = "ScanToBIM";
                application.CreateRibbonTab(tabName);

                // Création d'un panneau dans cet onglet
                RibbonPanel panel = application.CreateRibbonPanel(tabName, "Modélisation à partir d'un nuage de points");

                name = "btnOpenWindow"; // Nom technique du bouton
                text = "ScanToBIM"; // Texte affiché sur le bouton
                toolTip = "Modélisation à partir d'un nuage de points au format e57"; // Info-bulle du bouton (s'affiche quand la souris est sur le bouton)
                largeImageName = EmbeddedParentIconLocation + "bouton.png"; // Chemin de l'image de l'icône du bouton

                // Bouton pour ouvrir la fenêtre principale
                PushButtonData buttonData = new PushButtonData(
                    name,
                    text,
                    GetType().Assembly.Location,
                    "BimToMode.CommandStart")
                {
                    ToolTip = toolTip,
                    LargeImage = GetEmbeddedImage(largeImageName)
                };

                // Ajout du bouton au panneau
                PushButton? btnOpen = panel.AddItem(buttonData) as PushButton;

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur au démarrage de l'application : " + ex.Message, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private BitmapSource? GetEmbeddedImage(string name)
        {
            try
            {
                // Charge l'image depuis les ressources du projet
                Assembly a = Assembly.GetExecutingAssembly(); // Récupère l'assemblage (le projet en cours)
                Stream? s = a.GetManifestResourceStream(name); // Trouve l'image dans les ressources
                if (s == null)
                    return null;
                return BitmapFrame.Create(s); // Crée une image bitmap à partir du flux de données
            }
            catch
            {
                return null; // Si l'image n'est pas trouvée ou qu'il y a une erreur, retourne null
            }
        }
    }
}