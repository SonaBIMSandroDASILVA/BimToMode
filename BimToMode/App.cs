using Autodesk.Revit.UI;
using System;
using System.Windows;



namespace BimToMode
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Création d'un onglet personnalisé dans le ruban Revit
                string tabName = "BimToMode";
                application.CreateRibbonTab(tabName);

                // Création d'un panneau dans cet onglet
                RibbonPanel panel = application.CreateRibbonPanel(tabName, "Modélisation à partir d'un nuage de points");

                // Bouton pour ouvrir la fenêtre principale
                PushButtonData buttonData = new PushButtonData(
                    "btnOpenWindow", 
                    "ScanToBIM", 
                    GetType().Assembly.Location,
                    "BimToMode.CommandStart"
                );

                // Ajout du bouton au panneau
                panel.AddItem(buttonData);

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
    }
}
