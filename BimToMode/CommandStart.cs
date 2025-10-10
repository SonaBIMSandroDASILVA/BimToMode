using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Diagnostics;
using System.Windows;

namespace BimToMode
{
    // Attribut pour spécifier que cette commande est exécutée manuellement (pas automatiquement au démarrage)
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    internal class CommandStart : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Créer une instance de la fenêtre principale avec le document actif
                MainWindow mainWindow = new MainWindow(commandData.Application.ActiveUIDocument);

                mainWindow.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Impossible de charger la fenêtres", MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine("Add :" + "Impossible de charger la fenêtres" + ex.Message);
                return Result.Failed;
            }
        }
    }
}
