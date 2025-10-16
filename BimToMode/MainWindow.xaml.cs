using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;
using System.Windows.Controls;


namespace BimToMode
{
    public partial class MainWindow : Window
    {
        private UIDocument _uiDocument;
        private Document _document;
        private JToken _mursPorteursData;


        // Chemin fixe du JSON (écrit en dur)
        private readonly string _pythonOutputPath = @"U:\Documents\COURS\Projet\resultats_detectionv2.json";

        public MainWindow(UIDocument uiDocument)
        {
            InitializeComponent();
            _uiDocument = uiDocument;
            _document = _uiDocument.Document;

            // Désactiver le bouton Suivant au départ
            btnSuivant.IsEnabled = true;
        }

        private async void ExecutePythonButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnExecutePython.IsEnabled = false;
                btnExecutePython.Content = "⏳ Traitement en cours...";
                btnSuivant.IsEnabled = false;

                // Lancer le script Python et attendre la fin
                bool success = await RunPythonScriptAsync();

                success = true;

                if (success)
                {
                    MessageBox.Show(
                        "Le processus Python est terminé.\nVous pouvez maintenant cliquer sur 'Suivant'.",
                        "Succès", MessageBoxButton.OK, MessageBoxImage.Information);

                    // Activer le bouton Suivant
                    btnSuivant.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnExecutePython.IsEnabled = true;
                btnExecutePython.Content = "🔍 Lancer la Détection";
            }
        }

        private async Task<bool> RunPythonScriptAsync()
        {
            try
            {
                // Chemin de l'exécutable Python
                string exePath = @"C:\ProgramData\Autodesk\Revit\Addins\2026\BimToMode\PythonScripts\ScanToBIM.exe";

                if (!File.Exists(exePath))
                {
                    MessageBox.Show($"L'exécutable Python est introuvable :\n{exePath}",
                        "Fichier manquant", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }

                ProcessStartInfo start = new ProcessStartInfo
                {
                    FileName = exePath,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetDirectoryName(exePath)
                };

                using (Process process = new Process { StartInfo = start })
                {
                    process.Start();

                    // Attendre la fin du processus sans bloquer l'UI
                    await Task.Run(() => process.WaitForExit());

                    string errors = await process.StandardError.ReadToEndAsync();
                    if (!string.IsNullOrEmpty(errors))
                    {
                        MessageBox.Show($"Erreurs Python :\n{errors}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                        return false;
                    }

                    string output = await process.StandardOutput.ReadToEndAsync();
                    Debug.WriteLine($"Sortie Python : {output}");
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du lancement : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return true;
            }
        }

        private void SuivantButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnSuivant.IsEnabled = false;
                btnSuivant.Content = "⏳ Lecture du JSON...";

                ProcessPythonOutput();

                btnSuivant.Content = "✅ Résultats importés";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                btnSuivant.IsEnabled = true;
                btnSuivant.Content = "➡️ Suivant (Importer les résultats)";
            }
        }

        private void ProcessPythonOutput()
        {
            try
            {
                if (!File.Exists(_pythonOutputPath))
                    return;

                string jsonContent = File.ReadAllText(_pythonOutputPath);
                JObject data = JObject.Parse(jsonContent);

                // Nettoyer et récupérer Niveau 0
                Level levelZero = CleanAndGetLevelZero(_document);

                //
                bool isFirstLevelCreated = false;

                if (data["plans_horizontaux"] != null)
                {
                    var plans = data["plans_horizontaux"];

                    if (plans["sol"] != null)
                    {
                        CreateLevelsFromPlans(plans["sol"], "Sol", ref isFirstLevelCreated, levelZero);

                        // Créer les sols correspondants
                        CreateFloors(plans["sol"]);
                    }

                    if (plans["plafond"] != null)
                    {

                        // Créer les plafonds correspondants
                        CreateCeilings(plans["plafond"]);
                    }


                    //if (plans["toiture"] != null)
                    //{
                    //    CreateLevelsFromPlans(plans["toiture"], "Toiture", ref isFirstLevelCreated, levelZero);
                    //}

                }

                if (data["plans_verticaux"] != null)
                {
                    var murs = data["plans_verticaux"];

                    // Remplir la liste des types de murs dans l'UI
                    PopulateWallTypesSelection();

                    if (murs["mur_porteur"] != null)
                    {
                        _mursPorteursData = murs["mur_porteur"];


                    }

                    if (murs["cloison"] != null)
                    {
                        var selectedWallTypeIds = GetSelectedWallTypes();

                        // Créer les plafonds correspondants
                        ProcessMurs(murs["cloison"], selectedWallTypeIds);

                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lecture JSON : {ex.Message}");
            }
        }


        private void PopulateWallTypesSelection()
        {
            WallTypesStackPanel.Children.Clear(); // ✅ utiliser le StackPanel existant
            var wallTypes = new FilteredElementCollector(_document)
                                .OfClass(typeof(WallType))
                                .Cast<WallType>()
                                .OrderBy(wt => wt.Name);

            foreach (var wallType in wallTypes)
            {
                CheckBox cb = new CheckBox
                {
                    Content = wallType.Name,
                    Tag = wallType.Id,
                    IsChecked = false,
                    Margin = new Thickness(5)
                };
                cb.Checked += WallTypeCheckBox_Checked;
                cb.Unchecked += WallTypeCheckBox_Checked;

                WallTypesStackPanel.Children.Add(cb);
            }
        }


        private List<ElementId> GetSelectedWallTypes()
        {
            List<ElementId> selectedTypes = new List<ElementId>();

            foreach (var child in WallTypesStackPanel.Children)
            {
                if (child is CheckBox cb && cb.IsChecked == true && cb.Tag is ElementId id)
                {
                    selectedTypes.Add(id);
                }
            }

            return selectedTypes;
        }

        private void WallTypeCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            //btnSuivant.IsEnabled = WallTypesStackPanel.Children
            //    .OfType<CheckBox>()
            //    .Any(cb => cb.IsChecked == true);
        }



        /// <summary>
        ///Nettoie tous les niveaux sauf "Niveau 0" et retourne ce niveau. Si "Niveau 0" n'existe pas, il est créé.
        /// <summary>
        private Level CleanAndGetLevelZero(Document doc)
        {
            Level? levelZero = null;

            var levels = new FilteredElementCollector(doc)
                            .OfClass(typeof(Level))
                            .Cast<Level>()
                            .ToList();

            using (Transaction tx = new Transaction(doc, "Nettoyage des niveaux"))
            {
                tx.Start();

                // Chercher Niveau 0
                levelZero = levels.FirstOrDefault(l => l.Name.Equals("Niveau 0", StringComparison.OrdinalIgnoreCase));

                foreach (var lvl in levels)
                {
                    if (levelZero != null && lvl.Id == levelZero.Id)
                        continue;

                    // Supprime tous les autres niveaux
                    doc.Delete(lvl.Id);
                }

                // Créer Niveau 0 s’il n’existe pas
                if (levelZero == null)
                {
                    levelZero = Level.Create(doc, 0.0);
                    levelZero.Name = "Niveau 0";
                    Debug.WriteLine("✅ Niveau 0 créé");
                }

                tx.Commit();
            }

            return levelZero;
        }


        /// <summary>
        /// Crée ou met à jour les niveaux dans Revit à partir des plans d'une catégorie (sol/plafond/toiture)
        /// </summary>
        private void CreateLevelsFromPlans(JToken plansToken, string prefixName, ref bool isFirstLevelCreated, Level levelZero)
        {
            using (Transaction tx = new Transaction(_document, $"Création des niveaux {prefixName}"))
            {
                tx.Start();

                int index = 0;
                foreach (var plan in plansToken)
                {
                    double altitude = plan["altitude"]?.ToObject<double>() ?? 0.0;
                    double zFeet = Math.Round(altitude * 3.28084, 2);

                    Level lvl;

                    if (index == 0 && prefixName.Equals("Sol", StringComparison.OrdinalIgnoreCase))
                    {
                        // ⚡ Utiliser Niveau 0 existant
                        lvl = levelZero;

                        lvl.Elevation = zFeet;

                    }
                    else
                    {
                        lvl = CleanAndGetLevelAtElevation(_document, zFeet);
                    }

                    // Définir le nom unique
                    string baseName = prefixName.Equals("Sol", StringComparison.OrdinalIgnoreCase) ? $"Niveau {index}" : $"{prefixName} {index}";
                    string uniqueName = baseName;
                    int suffix = 1;
                    while (new FilteredElementCollector(_document)
                                .OfClass(typeof(Level))
                                .Cast<Level>()
                                .Any(l => l.Name.Equals(uniqueName, StringComparison.OrdinalIgnoreCase) && l.Id != lvl.Id))
                    {
                        uniqueName = $"{baseName}_{suffix}";
                        suffix++;
                    }
                    lvl.Name = uniqueName;

                    Debug.WriteLine($"✅ {uniqueName} créé ou mis à jour à Z={altitude:F3} m");
                    index++;
                }


                tx.Commit();
            }
        }



        /// <summary>
        /// 
        /// Crée des sols dans Revit à partir des données JSON
        /// 
        /// <summary>
        private void CreateFloors(JToken floorsToken)
        {
            // Récupérer un FloorType par défaut
            FloorType floorType = new FilteredElementCollector(_document)
                                    .OfClass(typeof(FloorType))
                                    .Cast<FloorType>()
                                    .FirstOrDefault(ft => ft.Name.Contains("Béton"))
                                ?? new FilteredElementCollector(_document)
                                    .OfClass(typeof(FloorType))
                                    .Cast<FloorType>()
                                    .First();

            using (Transaction tx = new Transaction(_document, "Création des sols"))
            {
                tx.Start();

                // Préparer la gestion des avertissements
                FailureHandlingOptions options = tx.GetFailureHandlingOptions();
                options.SetFailuresPreprocessor(new IgnoreWarnings());
                tx.SetFailureHandlingOptions(options);

                foreach (var floor in floorsToken)
                {
                    double altitudeMeters = floor["altitude"]?.ToObject<double>() ?? 0.0;
                    string name = floor["nom"]?.ToString() ?? "Sol";

                    // Convertir l'altitude en pieds pour Revit
                    double zFeet = Math.Round(altitudeMeters * 3.28084, 2);

                    Level level = CleanAndGetLevelAtElevation(_document, zFeet);

                    // Créer un CurveLoop à partir du contour JSON
                    CurveLoop loop = new CurveLoop();
                    var points = floor["contour"];
                    for (int i = 0; i < points.Count(); i++)
                    {
                        var p1 = points[i];
                        var p2 = points[(i + 1) % points.Count()];

                        // Conversion mètres → pieds et arrondi
                        double x1 = Math.Round(p1[0].ToObject<double>() * 3.28084, 3);
                        double y1 = Math.Round(p1[1].ToObject<double>() * 3.28084, 3);
                        double x2 = Math.Round(p2[0].ToObject<double>() * 3.28084, 3);
                        double y2 = Math.Round(p2[1].ToObject<double>() * 3.28084, 3);

                        XYZ start = new XYZ(x1, y1, zFeet);
                        XYZ end = new XYZ(x2, y2, zFeet);

                        loop.Append(Line.CreateBound(start, end));
                    }

                    // Création du sol
                    Floor newFloor = Floor.Create(_document, new List<CurveLoop> { loop }, floorType.Id, level.Id);
                    newFloor.Name = name;

                    Debug.WriteLine($"✅ Sol créé : {name} à Z={altitudeMeters:F3} m");
                }

                tx.Commit();
            }
        }



        /// <summary>
        /// Crée des sols dans Revit à partir des données JSON
        /// <summary>
        private void CreateCeilings(JToken ceilingsToken)
        {
            // Récupérer un type de plafond par défaut
            CeilingType? ceilingType = new FilteredElementCollector(_document)
                                        .OfClass(typeof(CeilingType))
                                        .Cast<CeilingType>()
                                        .FirstOrDefault();

            if (ceilingType == null)
            {
                MessageBox.Show("Aucun type de plafond trouvé dans le projet.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            using (Transaction tx = new Transaction(_document, "Création des plafonds"))
            {
                tx.Start();

                // Préparer la gestion des avertissements
                FailureHandlingOptions options = tx.GetFailureHandlingOptions();
                options.SetFailuresPreprocessor(new IgnoreWarnings());
                tx.SetFailureHandlingOptions(options);

                // Récupérer et trier tous les niveaux du projet
                var levels = new FilteredElementCollector(_document)
                                .OfClass(typeof(Level))
                                .Cast<Level>()
                                .OrderBy(l => l.Elevation)
                                .ToList();

                foreach (var ceiling in ceilingsToken)
                {
                    double altitudeMeters = Math.Round(ceiling["altitude"]?.ToObject<double>() ?? 0.0, 2);


                    // Trouver le niveau inférieur
                    Level lowerLevel = levels.LastOrDefault(l => l.Elevation <= altitudeMeters * 3.28084);
                    if (lowerLevel == null) lowerLevel = levels.First();

                    // Calculer l’altitude réelle en pieds
                    double lowerLevelMeters = lowerLevel.Elevation / 3.28084;
                    double altitudeReelleMeters = altitudeMeters + lowerLevelMeters;
                    double zFeet = Math.Round(altitudeReelleMeters * 3.28084, 2);

                    // Offset par rapport au niveau
                    double offsetFeet = zFeet - lowerLevel.Elevation;

                    var contour = ceiling["contour"];
                    if (contour == null || !contour.Any())
                    {
                        Debug.WriteLine($"⚠️ Plafond ignoré (aucun contour défini)");
                        continue;
                    }

                    // Créer le profil du plafond avec conversion mètres → pieds et arrondi
                    CurveLoop loop = new CurveLoop();
                    for (int i = 0; i < contour.Count(); i++)
                    {
                        var p1 = contour[i];
                        if (p1 == null || p1.Count() < 2) continue;
                        var p2 = contour[(i + 1) % contour.Count()];
                        if (p2 == null || p2.Count() < 2) continue;

                        double x1 = Math.Round(p1[0].ToObject<double>() * 3.28084, 3);
                        double y1 = Math.Round(p1[1].ToObject<double>() * 3.28084, 3);
                        double x2 = Math.Round(p2[0].ToObject<double>() * 3.28084, 3);
                        double y2 = Math.Round(p2[1].ToObject<double>() * 3.28084, 3);

                        XYZ start = new XYZ(x1, y1, zFeet);
                        XYZ end = new XYZ(x2, y2, zFeet);

                        loop.Append(Line.CreateBound(start, end));
                    }

                    try
                    {
                        List<CurveLoop> profile = new List<CurveLoop> { loop };

                        // Créer le plafond attaché au niveau
                        Ceiling newCeiling = Ceiling.Create(_document, profile, ceilingType.Id, lowerLevel.Id);

                        // Appliquer l’offset vertical
                        Parameter offsetParam = newCeiling.get_Parameter(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM);
                        if (offsetParam != null && !offsetParam.IsReadOnly)
                            offsetParam.Set(offsetFeet);

                        // Nom unique
                        string baseName = $"Plafond_{altitudeReelleMeters:F2}m";
                        string uniqueName = baseName;
                        int suffix = 1;
                        while (new FilteredElementCollector(_document)
                                    .OfClass(typeof(Ceiling))
                                    .Cast<Ceiling>()
                                    .Any(c => c.Name.Equals(uniqueName, StringComparison.OrdinalIgnoreCase)))
                        {
                            uniqueName = $"{baseName}_{suffix}";
                            suffix++;
                        }

                        newCeiling.Name = uniqueName;
                        Debug.WriteLine($"✅ Plafond créé : {uniqueName} (niveau {lowerLevel.Name}, offset={offsetFeet:F3} ft)");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"❌ Erreur création plafond à Z={altitudeMeters:F3} m: {ex.Message}");
                    }
                }

                tx.Commit();
            }
        }

        /// <summary>
        /// Fonction utilitaire pour récupérer ou créer un Level à une altitude donnée
        /// <summary>
        private Level CleanAndGetLevelAtElevation(Document doc, double elevation)
        {
            var levels = new FilteredElementCollector(doc)
                            .OfClass(typeof(Level))
                            .Cast<Level>()
                            .ToList();

            Level? lvl = levels.FirstOrDefault(l => Math.Abs(l.Elevation - elevation) < 0.01);
            if (lvl == null)
            {
                // ⚠️ Ici on ne crée pas de transaction car l'appelant a déjà une transaction ouverte
                lvl = Level.Create(doc, elevation);
                // Le nom sera attribué par CreateLevelsFromPlans
            }
            return lvl;
        }





        /// <summary>
        /// crée des murs dans Revit à partir des données JSON
        /// <summary
        private void ProcessMurs(JToken mursToken, List<ElementId> selectedWallTypes)
        {
            if (mursToken == null || !mursToken.Any())
            {
                Debug.WriteLine("Aucun mur à créer.");
                return;
            }

            // Récupérer tous les niveaux du document
            var levels = new FilteredElementCollector(_document)
                            .OfClass(typeof(Level))
                            .Cast<Level>()
                            .OrderBy(l => l.Elevation)
                            .ToList();

            // Récupérer les WallType à partir des ElementId sélectionnés
            var wallTypes = selectedWallTypes
                            .Select(id => _document.GetElement(id) as WallType)
                            .Where(wt => wt != null)
                            .ToList();

            if (wallTypes.Count == 0)
            {
                MessageBox.Show("Aucun type de mur sélectionné.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            using (Transaction tx = new Transaction(_document, "Création des murs depuis JSON"))
            {
                tx.Start();

                int wallIndex = 1;

                foreach (var mur in mursToken)
                {
                    string nom = mur["id"]?.ToString() ?? $"Mur {wallIndex}";
                    double epaisseurJSON = Math.Round(mur["epaisseur_m"]?.ToObject<double>() ?? 0.15, 2);

                    var startPointToken = mur["start"];
                    var endPointToken = mur["end"];
                    if (startPointToken == null || endPointToken == null)
                        continue;

                    Level? levelZero = GetLevelZero(_document);
                    if (levelZero == null)
                    {
                        MessageBox.Show("⚠ Aucun Niveau 0 trouvé dans le projet.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // Conversion mètres -> pieds arrondie
                    double xStart = Math.Round(startPointToken[0].ToObject<double>() * 3.28084, 2);
                    double yStart = Math.Round(startPointToken[1].ToObject<double>() * 3.28084, 2);
                    double xEnd = Math.Round(endPointToken[0].ToObject<double>() * 3.28084, 2);
                    double yEnd = Math.Round(endPointToken[1].ToObject<double>() * 3.28084, 2);

                    XYZ start = new XYZ(xStart, yStart, levelZero.Elevation);
                    XYZ end = new XYZ(xEnd, yEnd, levelZero.Elevation);

                    // 🔹 Choisir le type de mur dont l'épaisseur ≤ JSON et la plus proche
                    // 🔹 Sélectionner le type de mur dont l'épaisseur est la plus proche de la valeur JSON
                    WallType? wallTypeToUse = wallTypes
                        .OrderBy(wt => Math.Abs((wt.Width * 0.3048) - epaisseurJSON)) // conversion en mètres
                        .FirstOrDefault();

                    // 🔸 Sécurité : si aucun type trouvé, on prend le plus mince
                    if (wallTypeToUse == null)
                        wallTypeToUse = wallTypes.OrderBy(wt => wt.Width).First();



                    // Créer le mur attaché au Niveau 0
                    Wall wall = Wall.Create(_document, Line.CreateBound(start, end), wallTypeToUse.Id, levelZero.Id, 10.0, 0.0, false, false);

                    // Définir l'épaisseur si le paramètre est modifiable
                    Parameter widthParam = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                    if (widthParam != null && !widthParam.IsReadOnly)
                        widthParam.Set(epaisseurJSON * 304.8); // m -> mm

                    // Gestion de la contrainte supérieure
                    Level? upperLevel = levels.FirstOrDefault(l => l.Elevation > levelZero.Elevation);
                    if (upperLevel != null)
                    {
                        Parameter topConstraint = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                        if (topConstraint != null && !topConstraint.IsReadOnly)
                            topConstraint.Set(upperLevel.Id);

                        Parameter unconnectedHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                        if (unconnectedHeight != null && !unconnectedHeight.IsReadOnly)
                            unconnectedHeight.Set(0.0);
                    }

                    Debug.WriteLine($"✅ Mur créé : Nom={nom}, ID={wall.Id}, Type={wallTypeToUse.Name}, Épaisseur={epaisseurJSON} m, Start=({start.X:F2},{start.Y:F2}), End=({end.X:F2},{end.Y:F2}), Niveau Supérieur={(upperLevel != null ? upperLevel.Name : "Sans")}");
                    wallIndex++;
                }

                tx.Commit();
            }
        }





        /// <summary>
        /// Récupère le Niveau 0 existant dans le document.
        /// Retourne null si aucun Niveau 0 n'existe.
        /// </summary>
        private Level? GetLevelZero(Document doc)
        {
            return new FilteredElementCollector(doc)
                        .OfClass(typeof(Level))
                        .Cast<Level>()
                        .FirstOrDefault(l => l.Name.Equals("Niveau 0", StringComparison.OrdinalIgnoreCase));
        }



        /// <summary>
        /// Pose des ouvertures (portes/fenêtres) dans les murs à partir des données JSON
        /// </summary>
        /// <param name="suppressOnError">Si true, supprime l'ouverture en cas d'erreur de coupe</param>
        private void ProcessOuverture(JToken ouverturesToken)
        {
            if (ouverturesToken == null)
            {
                MessageBox.Show("Aucune ouverture à créer.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            using (Transaction tx = new Transaction(_document, "Création des ouvertures"))
            {
                tx.Start();

                // 🔹 Récupérer le FamilySymbol "Percement_parois"
                FamilySymbol percSymbol = new FilteredElementCollector(_document)
                                            .OfClass(typeof(FamilySymbol))
                                            .Cast<FamilySymbol>()
                                            .FirstOrDefault(f => f.Name == "Percement_parois");

                if (percSymbol == null)
                {
                    MessageBox.Show("Famille 'Percement_parois' non trouvée.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (!percSymbol.IsActive)
                    percSymbol.Activate();

                JObject ouverturesObj = ouverturesToken as JObject;
                if (ouverturesObj == null)
                {
                    MessageBox.Show("Format JSON invalide pour les ouvertures.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                JArray portesToken = ouverturesObj["porte"] as JArray;
                if (portesToken == null || portesToken.Count == 0)
                {
                    MessageBox.Show("Aucune ouverture à créer.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // 🔹 Récupérer tous les murs du document
                var murs = new FilteredElementCollector(_document)
                            .OfClass(typeof(Wall))
                            .Cast<Wall>()
                            .ToList();

                foreach (var porte in portesToken)
                {
                    JArray centre = porte["centre_3d"] as JArray;
                    if (centre == null || centre.Count < 3) continue;

                    double x = Math.Round(centre[0].ToObject<double>() * 3.28084, 2);
                    double y = Math.Round(centre[1].ToObject<double>() * 3.28084, 2);
                    double z = Math.Round(centre[2].ToObject<double>() * 3.28084, 2);
                    XYZ point = new XYZ(x, y, z);

                    Wall nearestWall = murs
                        .OrderBy(w =>
                        {
                            if (w.Location is LocationCurve lc)
                            {
                                Curve c = lc.Curve;
                                IntersectionResult ir = c.Project(point);
                                if (ir != null)
                                    return (ir.XYZPoint - point).GetLength();
                            }
                            return double.MaxValue;
                        })
                        .FirstOrDefault();

                    if (nearestWall == null)
                    {
                        Debug.WriteLine("⚠ Aucun mur trouvé à proximité du point");
                        continue;
                    }

                    try
                    {
                        FamilyInstance percInstance = _document.Create.NewFamilyInstance(
                            point,
                            percSymbol,
                            nearestWall,
                            Autodesk.Revit.DB.Structure.StructuralType.NonStructural
                        );

                        // 🔹 Mettre à jour les paramètres
                        Parameter largeurParam = percInstance.LookupParameter("Largeur");
                        if (largeurParam != null && !largeurParam.IsReadOnly)
                            largeurParam.Set(Math.Round((porte["largeur"]?.ToObject<double>() ?? 1.0) * 3.28084, 2));

                        Parameter hauteurParam = percInstance.LookupParameter("Hauteur");
                        if (hauteurParam != null && !hauteurParam.IsReadOnly)
                            hauteurParam.Set(Math.Round((porte["hauteur"]?.ToObject<double>() ?? 2.0) * 3.28084, 2));

                        Parameter seuilParam = percInstance.LookupParameter("AltitudeSeuil");
                        if (seuilParam != null && !seuilParam.IsReadOnly)
                            seuilParam.Set(Math.Round((porte["altitude_seuil"]?.ToObject<double>() ?? 0.0) * 3.28084, 2));

                        Parameter Decalagehote = percInstance.LookupParameter("Décalage par rapport à l'hôte");
                        if (Decalagehote != null && !Decalagehote.IsReadOnly)
                            Decalagehote.Set(0.0);

                        // 🔹 Appliquer la rotation selon la normale uniquement
                        JToken orientationToken = porte["orientation"];
                        if (orientationToken != null)
                        {
                            JArray normaleArray = orientationToken["normale"] as JArray;
                            XYZ normale = XYZ.BasisX;

                            if (normaleArray != null && normaleArray.Count == 3)
                            {
                                // Arrondir à -1, 0 ou 1 puis mettre en valeur absolue
                                double nx = Math.Abs(Math.Round(normaleArray[0].ToObject<double>()));
                                double ny = Math.Abs(Math.Round(normaleArray[1].ToObject<double>()));
                                normale = new XYZ(nx, ny, 0);

                                // Normale verticale → utiliser X+
                                if (normale.IsZeroLength())
                                    normale = new XYZ(1, 0, 0);
                            }


                            // Rotation autour de l’axe vertical
                            Line rotationAxis = Line.CreateBound(point, point + XYZ.BasisZ);
                            double angle = Math.Atan2(normale.Y, normale.X);

                            ElementTransformUtils.RotateElement(_document, percInstance.Id, rotationAxis, angle);
                        }

                        try
                        {
                            InstanceVoidCutUtils.AddInstanceVoidCut(_document, nearestWall, percInstance);
                            Debug.WriteLine($"✅ Instance posée à ({x},{y},{z}) coupe le mur {nearestWall.Id}");
                        }
                        catch (Exception exCut)
                        {
                            Debug.WriteLine($"⚠ Erreur lors de la coupe du mur {nearestWall.Id}: {exCut.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"❌ Erreur création ouverture: {ex.Message}");
                    }
                }

                tx.Commit();
            }

            MessageBox.Show("✅ Ouvertures créées avec succès.", "Succès", MessageBoxButton.OK, MessageBoxImage.Information);
        }




        /// <summary>
        /// Création de coupes centrées dans le modèle
        /// <summary
        private void Créationcoupe(Document doc)
        {
            try
            {
                // 🔹 Récupération de tous les murs du projet
                var murs = new FilteredElementCollector(doc)
                    .OfClass(typeof(Wall))
                    .Cast<Wall>()
                    .ToList();

                if (murs.Count == 0)
                {
                    MessageBox.Show("Aucun mur trouvé pour créer les coupes.");
                    return;
                }

                // 🔹 Calcul du rectangle englobant tous les murs
                double xMin = double.MaxValue, yMin = double.MaxValue, xMax = double.MinValue, yMax = double.MinValue;
                double zMin = double.MaxValue, zMax = double.MinValue;

                foreach (var mur in murs)
                {
                    LocationCurve lc = mur.Location as LocationCurve;
                    if (lc == null) continue;

                    XYZ p1 = lc.Curve.GetEndPoint(0);
                    XYZ p2 = lc.Curve.GetEndPoint(1);

                    xMin = Math.Min(xMin, Math.Min(p1.X, p2.X));
                    yMin = Math.Min(yMin, Math.Min(p1.Y, p2.Y));
                    xMax = Math.Max(xMax, Math.Max(p1.X, p2.X));
                    yMax = Math.Max(yMax, Math.Max(p1.Y, p2.Y));

                    zMin = Math.Min(zMin, Math.Min(p1.Z, p2.Z));
                    zMax = Math.Max(zMax, Math.Max(p1.Z, p2.Z));
                }

                // 🔹 Calcul du centre du rectangle (milieu du L et du l)
                double xCenter = (xMin + xMax) / 2;
                double yCenter = (yMin + yMax) / 2;

                // 🔹 Hauteurs configurables (en pieds)
                double offsetBas = -3.28084;   // -1 m sous le niveau bas
                double offsetHaut = 6.56168;   // +2 m au-dessus du plus haut point

                double bottom = zMin + offsetBas;
                double top = zMax + offsetHaut;

                using (Transaction tx = new Transaction(doc, "Création des coupes"))
                {
                    tx.Start();

                    // --- Coupe dans le sens X (vue sur Y)
                    CreateSectionView(doc,
                        new XYZ(xCenter, yMin - 5, 0), // ligne centrale dans le sens X
                        new XYZ(xCenter, yMax + 5, 0),
                        bottom, top,
                        "Coupe_X_Centrale");

                    // --- Coupe dans le sens Y (vue sur X)
                    CreateSectionView(doc,
                        new XYZ(xMin - 5, yCenter, 0),
                        new XYZ(xMax + 5, yCenter, 0),
                        bottom, top,
                        "Coupe_Y_Centrale");

                    tx.Commit();
                }

                Debug.WriteLine("✅ Coupes centrales créées avec succès !");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Erreur lors de la création des coupes : {ex.Message}");
            }
        }

        /// <summary>
        /// Crée une vue de section entre deux points avec une hauteur définie
        /// <summary
        private void CreateSectionView(Document doc, XYZ p1, XYZ p2, double zMin, double zMax, string name)
        {
            ViewFamilyType viewFamilyType = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Section);

            if (viewFamilyType == null)
                throw new Exception("Aucun type de vue de section trouvé.");

            // 🔹 Direction de la coupe
            XYZ direction = (p2 - p1).Normalize();
            XYZ up = XYZ.BasisZ;
            XYZ viewDir = direction.CrossProduct(up);

            // 🔹 Origine de la coupe au centre
            XYZ origin = (p1 + p2) / 2;

            // 🔹 Définition de la boîte de coupe
            BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
            sectionBox.Enabled = true;

            double halfLength = p1.DistanceTo(p2) / 2;
            sectionBox.Min = new XYZ(-halfLength, zMin - origin.Z, -5);
            sectionBox.Max = new XYZ(halfLength, zMax - origin.Z, 5);

            // 🔹 Orientation de la coupe
            Transform transform = Transform.Identity;
            transform.Origin = origin;
            transform.BasisX = direction; // Axe du mur (horizontale)
            transform.BasisY = up;        // Vertical
            transform.BasisZ = viewDir;   // Profondeur (vue)
            sectionBox.Transform = transform;

            // 🔹 Création de la coupe
            ViewSection viewSection = ViewSection.CreateSection(doc, viewFamilyType.Id, sectionBox);
            viewSection.Name = name;
        }

        /// <summary>
        /// Méthode appelée lorsque l'utilisateur clique sur "✅ Valider les murs"
        /// </summary>
        private void ValidateWallsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedWallTypes = GetSelectedWallTypes();

                if (selectedWallTypes == null || selectedWallTypes.Count == 0)
                {
                    MessageBox.Show("Veuillez sélectionner au moins un type de mur avant de valider.",
                                    "Aucun mur sélectionné", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (_mursPorteursData == null)
                {
                    MessageBox.Show("Aucune donnée de mur porteur n’a été chargée.",
                                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Créer les murs
                ProcessMurs(_mursPorteursData, selectedWallTypes);

                // Désactiver les cases et le bouton pour éviter un double traitement
                foreach (var child in WallTypesStackPanel.Children.OfType<CheckBox>())
                    child.IsEnabled = false;

                btnValidateWalls.IsEnabled = false;
                btnSuivant.IsEnabled = false;

                MessageBox.Show("✅ Les murs ont été créés avec succès.",
                                "Succès", MessageBoxButton.OK, MessageBoxImage.Information);

                // 🔹 Ici, poursuivre automatiquement sur les percements
                ProcessOuverture(JObject.Parse(File.ReadAllText(_pythonOutputPath))["ouvertures"]);

                // Puis sur les coupes
                Créationcoupe(_document);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la création des murs : {ex.Message}",
                                "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// module de gestion des avertissements Revit
        /// </summary>
        public class IgnoreWarnings : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                foreach (FailureMessageAccessor f in failuresAccessor.GetFailureMessages())
                {
                    failuresAccessor.DeleteWarning(f);
                }
                return FailureProcessingResult.Continue;
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
