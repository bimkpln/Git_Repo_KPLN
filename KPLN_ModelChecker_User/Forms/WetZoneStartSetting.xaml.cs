using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.ExternalCommands;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;


namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Логика взаимодействия для WetZoneStartSetting.xaml
    /// </summary>
    public partial class WetZoneReviewWindow : Window
    {
        public static UIDocument _uiDoc;
        public List<Element> allRooms;
        public List<string> _invalidEquipment;
        public string _selectedParam;

        public WetZoneReviewWindow(UIDocument uiDoc, Document doc, string selectedParam, List<Element> livingRooms, List<Element> kitchenRooms, List<Element> wetRooms, List<Element> undefinedRooms)
        {
            InitializeComponent();
            _uiDoc = uiDoc;
            _selectedParam = selectedParam;

            LivingExp.Header = $"Жилые комнаты ({livingRooms.Count})";
            WetExp.Header = $"Мокрые зоны ({wetRooms.Count})";
            KitchenExp.Header = $"Кухни ({kitchenRooms.Count})";
            UndefinedExp.Header = $"Необработанные помещения помещения ({undefinedRooms.Count})";

            LivingList.ItemsSource = FormatRooms(livingRooms);
            WetList.ItemsSource = FormatRooms(wetRooms);
            KitchenList.ItemsSource = FormatRooms(kitchenRooms);           
            UndefinedList.ItemsSource = FormatRooms(undefinedRooms);

            _invalidEquipment = WetZoneCategories.InvalidEquipment.Any() ? WetZoneCategories.InvalidEquipment : null;

            BuildInfoReport(doc, livingRooms, kitchenRooms, wetRooms, undefinedRooms);        
        }

        /// <summary>
        /// Анализ документа на наличие необходимых условий + вывод отчёта
        /// </summary>
        public void BuildInfoReport(Document doc, List<Element> livingRooms, List<Element> kitchenRooms, List<Element> wetRooms, List<Element> undefinedRooms)
        {
            Paragraph report = new Paragraph();
            report.Inlines.Add(new Run($"Анализ. {doc.Title}:")
            {
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.Black,
                FontSize = 14
            });
            report.Inlines.Add(new LineBreak());
            report.Inlines.Add(new LineBreak());

            bool hasCriticalLevelError = false;
            bool hasCriticalParamError = false;

            // Проверка помещений
            allRooms = livingRooms.Concat(kitchenRooms).Concat(wetRooms).ToList();
            if (allRooms.Count == 0)
            {
                report.Inlines.Add(new Run("❎ Нет доступных элементов в категориях 'Жилые комнаты', 'Мокрые зоны', 'Кухни'.\n")
                {
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.Red
                });

                hasCriticalLevelError = true;
                hasCriticalParamError = true;
            }
            else
            {
                report.Inlines.Add(new Run("✅ Элементы распределены по категориям.\n")
                {
                   
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.DarkGreen
                });
            }
            report.Inlines.Add(new LineBreak());

            // Проверяем уровни привязки элементов
            Dictionary<Element, string> invalidLevelElements = new Dictionary<Element, string>();

            List<Inline> CheckLevels()
            {
                Dictionary<string, string> results = new Dictionary<string, string>();
                
                foreach (Element el in allRooms)
                {
                    ElementId levelId = el.LevelId;

                    if (levelId == ElementId.InvalidElementId)
                    {
                        invalidLevelElements[el] = el.Name ?? "<?>";
                        continue;
                    }

                    Level level = el.Document.GetElement(levelId) as Level;
                    if (level == null) continue; 
                    
                    string levelName = level.Name;

                    if (results.ContainsKey(levelName)) continue;

                    string[] parts = levelName.Split('_');
                    string parsed = null;

                    if (parts.Length >= 1 && int.TryParse(parts[0], out int first))
                    {
                        parsed = first.ToString();
                    }
                    else if (parts.Length >= 2 && int.TryParse(parts[1], out int second))
                    {
                        parsed = second.ToString();
                    }

                    if (parsed != null)
                    {
                        results[levelName] = parsed;
                    }
                    else
                    {
                        results[levelName] = "ОШИБКА. Имя не соответствует BEP.";
                        hasCriticalLevelError = true;
                    }
                }

                List<Inline> output = new List<Inline>();

                output.Add(new Run("🔎 Результаты преобразования названия уровней в этажи:\n") { FontWeight = FontWeights.SemiBold });

                int maxToShow = 30;
                int count = 0;

                foreach (var pair in results)
                {
                    if (count >= maxToShow) break;

                    bool isError = pair.Value == "ОШИБКА. Имя не соответствует BEP.";
                    output.Add(new Run($"       • {pair.Key} → {pair.Value}\n")
                    {
                        Foreground = isError ? Brushes.IndianRed : Brushes.DarkGreen
                    });

                    count++;
                }

                if (results.Count > maxToShow)
                {
                    int hiddenCount = results.Count - maxToShow;
                    output.Add(new Run($"       А также другие уровни ({hiddenCount}), которые не удалось обработать.\n")
                    {
                        Foreground = Brushes.IndianRed
                    });
                }

                return output;
            }

            List<Inline> levelReport = CheckLevels();
            foreach (var inline in levelReport)
            {
                report.Inlines.Add(inline);
            }

            // Элементы не привязанные к уровню
            List<Inline> outputInvalidLevelElements = new List<Inline>();
            if (invalidLevelElements.Count > 0)
            {
                outputInvalidLevelElements.Add(new LineBreak());
                outputInvalidLevelElements.Add(new Run("⚠️ Помещения без привязки к уровню:\n")
                {
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.Red
                });

                int maxToShow = 5;
                int count = 0;

                foreach (var pair in invalidLevelElements)
                {
                    if (count >= maxToShow) break;

                    string display = $"       • {pair.Value} [ID {pair.Key.Id.IntegerValue}]\n";
                    outputInvalidLevelElements.Add(new Run(display)
                    {
                        Foreground = Brushes.IndianRed
                    });

                    count++;
                }

                if (invalidLevelElements.Count > maxToShow)
                {
                    int remaining = invalidLevelElements.Count - maxToShow;
                    outputInvalidLevelElements.Add(new Run($"       А также другие помещения ({remaining}), которые не привязаны к уровню.\n")
                    {
                        Foreground = Brushes.IndianRed
                    });
                }

                foreach (var inline in outputInvalidLevelElements)
                {
                    report.Inlines.Add(inline);
                }

                hasCriticalLevelError = true;
            }

            report.Inlines.Add(new LineBreak());

            // Заполненость параметра ПОМ_Номер этажа
            List<Element> missingPNE = allRooms.Where(el =>
            {
                var param = el.LookupParameter("ПОМ_Номер этажа");
                return param == null || string.IsNullOrWhiteSpace(param.AsString());
            }).ToList();

            if (missingPNE.Count > 0)
            {
                report.Inlines.Add(new Run("⚠️ Внимание: у следующих помещений не заполнен параметр 'ПОМ_Номер этажа':\n")
                {
                    FontWeight = FontWeights.SemiBold,
                    Foreground = Brushes.Red
                });

                int maxToShow = 5;
                int count = 0;

                foreach (Element el in missingPNE)
                {
                    if (count >= maxToShow) break;

                    string name = el.Name ?? "<без имени>";
                    string id = el.Id.IntegerValue.ToString();

                    report.Inlines.Add(new Run($"       • {name} [ID {id}]\n")
                    {
                        Foreground = Brushes.IndianRed
                    });

                    count++;
                }

                if (missingPNE.Count > maxToShow)
                {
                    int remaining = missingPNE.Count - maxToShow;
                    report.Inlines.Add(new Run($"       А также ещё {remaining} помещений не содержат параметр 'ПОМ_Номер этажа'\n")
                    {
                        Foreground = Brushes.IndianRed
                    });
                }
                hasCriticalParamError = true;
            }
            else
            {
                report.Inlines.Add(new Run("✅ Все помещения содержат параметр 'ПОМ_Номер этажа'.\n")
                {
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.DarkGreen
                });
            }

            report.Inlines.Add(new LineBreak());

            // Заполненость параметра КВ_Номер
            List<Element> missingKVN = allRooms.Where(el =>
            {
                var param = el.LookupParameter("КВ_Номер");
                return param == null || string.IsNullOrWhiteSpace(param.AsString());
            }).ToList();
         
            if (missingKVN.Count > 0)
            {
                report.Inlines.Add(new Run("⚠️ Внимание: у следующих помещений не заполнен параметр 'КВ_Номер':\n")
                {
                    FontWeight = FontWeights.SemiBold,
                    Foreground = Brushes.Red
                });

                int maxToShow = 10;
                int count = 0;

                foreach (Element el in missingKVN)
                {
                    if (count >= maxToShow) break;

                    string name = el.Name ?? "<без имени>";
                    string id = el.Id.IntegerValue.ToString();

                    report.Inlines.Add(new Run($"       • {name} [ID {id}]\n")
                    {
                        Foreground = Brushes.IndianRed
                    });

                    count++;
                }

                if (missingKVN.Count > maxToShow)
                {
                    int remaining = missingKVN.Count - maxToShow;
                    report.Inlines.Add(new Run($"       А также ещё {remaining} помещений не содержат параметр 'КВ_Номер'\n")
                    {
                        Foreground = Brushes.IndianRed
                    });
                }

                hasCriticalLevelError = true;
                hasCriticalParamError = true;
            }
            else
            {
                report.Inlines.Add(new Run("✅ Все помещения содержат параметр 'КВ_Номер'.\n")
                {
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.DarkGreen
                });
            }

            report.Inlines.Add(new LineBreak());
       
            if (hasCriticalLevelError)
            {
               report.Inlines.Add(new Run("⛔ Плагин не может продолжить работу в режиме 'Уровень'.\n")
                {
                    Foreground = Brushes.Red,
                    FontWeight = FontWeights.Bold
                });
               Start1Button.IsEnabled = false;
           }

            if (hasCriticalParamError)
            {
                report.Inlines.Add(new Run("⛔ Плагин не может продолжить работу в режиме 'ПОМ_Номер этажа'.\n")
                {
                    Foreground = Brushes.Red,
                    FontWeight = FontWeights.Bold
                });
                Start2Button.IsEnabled = false;
            }
             
            if (undefinedRooms.Count > 0)
            {
                report.Inlines.Add(new Run("⛔ Плагин не может продолжить работу. Имеются необработанные помещения.\n      Для решения данной проблемы обратитесь в BIM-отдел.")
                {
                    Foreground = Brushes.Red,
                    FontWeight = FontWeights.Bold
                });

                Start1Button.IsEnabled = false;
                Start2Button.IsEnabled = false;
            }

            if (!hasCriticalLevelError && !hasCriticalParamError && undefinedRooms.Count == 0)
            {
                report.Inlines.Add(new Run("✅ Все проверки пройдены.")
                {
                    Foreground = Brushes.DarkGreen,
                    FontWeight = FontWeights.Bold
                });
                Start1Button.IsEnabled = true;
                Start2Button.IsEnabled = true;
            }
            else if (hasCriticalLevelError && !hasCriticalParamError && undefinedRooms.Count == 0)
            {
                report.Inlines.Add(new Run("⚠️ Плагин продолжит работу в режиме `Запуск (ПОМНомер этажа)`.")
                {
                    Foreground = Brushes.BlueViolet,
                    FontWeight = FontWeights.Bold
                });
                Start1Button.IsEnabled = false;
                Start2Button.IsEnabled = true;
            }
            else if (!hasCriticalLevelError && hasCriticalParamError && undefinedRooms.Count == 0)
            {
                report.Inlines.Add(new Run("⚠️ Плагин продолжит работу в режиме `Запуск (Уровень)`.")
                {
                    Foreground = Brushes.BlueViolet,
                    FontWeight = FontWeights.Bold
                });
                Start1Button.IsEnabled = true;
                Start2Button.IsEnabled = false;
            }

            InfoText.Document.Blocks.Clear();
            InfoText.Document.Blocks.Add(report);
        }
        
        // Формирование списка помещений внутри категорий
        private List<string> FormatRooms(List<Element> rooms)
        {
            return rooms
                .OrderBy(room => room.LookupParameter("КВ_Номер")?.AsString())
                .Select(room =>
                {
                    string name = room.Name ?? "<Без имени>";
                    string id = room.Id.IntegerValue.ToString();
                    string kv = room.LookupParameter("КВ_Номер")?.AsString() ?? "-";
                    return $"КВ_Номер: {kv} - {name} ({id})";
                }).ToList();
        }






        // XAML. Уровень
        private void Start1Button_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<int, List<Element>> roomsByLevel = allRooms.Where(r => r.LevelId != ElementId.InvalidElementId).Select(r =>
            {
                Level level = r.Document.GetElement(r.LevelId) as Level;
                return new { Room = r, Level = level };
            }).Where(x => x.Level != null).Select(x => {
                    string[] parts = x.Level.Name.Split('_');
                    if (parts.Length >= 1 && int.TryParse(parts[0], out int first))
                        return new { x.Room, LevelNumber = first };
                    else if (parts.Length >= 2 && int.TryParse(parts[1], out int second))
                        return new { x.Room, LevelNumber = second };
                    else
                        return null; 
            }).Where(x => x != null).GroupBy(x => x.LevelNumber).ToDictionary(g => g.Key, g => g.Select(x => x.Room).ToList());

            Dictionary<int, List<FamilyInstance>> familyInstancesByLevel = new Dictionary<int, List<FamilyInstance>>();
            IList<Element> allElements = new FilteredElementCollector(_uiDoc.Document).WhereElementIsNotElementType().ToElements();
            foreach (var element in allElements)
            {
                if (element is FamilyInstance familyInstance)
                {
                    string familyName = familyInstance.Symbol.Family.Name;


                    if (WetZoneCategories.InvalidEquipment.Contains(familyName))
                    {
                        Level level = _uiDoc.Document.GetElement(familyInstance.LevelId) as Level;
                        if (level != null)
                        {
                            string[] parts = level.Name.Split('_');
                            int levelNumber = 999;

                            if (parts.Length >= 1 && int.TryParse(parts[0], out levelNumber)) { }   
                            else if (parts.Length >= 2 && int.TryParse(parts[1], out levelNumber)) { }

                            if (!familyInstancesByLevel.ContainsKey(levelNumber))
                            {
                                familyInstancesByLevel[levelNumber] = new List<FamilyInstance>();
                            }
                            familyInstancesByLevel[levelNumber].Add(familyInstance);
                        }
                    }
                }
            }
            Dictionary<int, List<FamilyInstance>> sortedFamilyInstancesByLevel = familyInstancesByLevel.OrderBy(kvp => kvp.Key).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            CheckWetZoneViolations(roomsByLevel, sortedFamilyInstancesByLevel, _selectedParam);
            this.Close();
        }

        // XAML. ПОМ_Номер этажа
        private void Start2Button_Click(object sender, RoutedEventArgs e)
        {
            TaskDialog.Show("Предупреждение", 
                "В данном режиме стиральные машины и другое оборудование, требующее подключения к водопроводным сетям или являющегося источником шума, вибраций, не учитывается.");

            Dictionary<int, List<Element>> roomsByFloorParam = allRooms.Select(r =>
            {
                Parameter param = r.LookupParameter("ПОМ_Номер этажа");
                if (param == null || !param.HasValue) return null;

                string val = param.AsString();
                if (string.IsNullOrWhiteSpace(val)) return null;

                if (int.TryParse(val, out int floorNumber))
                    return new { Room = r, Floor = floorNumber };
                else
                    return null;
            }).Where(x => x != null).GroupBy(x => x.Floor).ToDictionary(g => g.Key, g => g.Select(x => x.Room).ToList());

           CheckWetZoneViolations(roomsByFloorParam, null, _selectedParam);
            this.Close();
        }

        // XAML. Отмена
        private void ButtonCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }











        // Метод поиска помещений согласно условиям
        public static void CheckWetZoneViolations(Dictionary<int, List<Element>> roomsByFloorParam, Dictionary<int, List<FamilyInstance>> familyInstancesByLevel, string _selectedParam)
        {
            List<List<Element>> KitchenOverLiving_Illegal = new List<List<Element>>();
            List<List<Element>> WetOverLiving_Illegal = new List<List<Element>>();    
            
            List<List<Element>> KitchenUnderWet_Illegal = new List<List<Element>>();
            List<List<Element>> KitchenUnderWet_Accepted = new List<List<Element>>();

            List<List<Element>> InvalidEquipmentOverLiving_Illegal = new List<List<Element>>();

            List<int> orderedFloors = roomsByFloorParam.Keys.OrderBy(f => f).ToList();
            foreach (int floor in orderedFloors)
            {
                int upperFloor = floor + 1;
                if (!roomsByFloorParam.ContainsKey(upperFloor)) continue;

                List<Room> lowerRooms = roomsByFloorParam[floor].OfType<Room>().ToList();
                List<Room> upperRooms = roomsByFloorParam[upperFloor].OfType<Room>().ToList();

                foreach (var lower in lowerRooms)
                {
                    string lowerType = lower.LookupParameter(_selectedParam)?.AsString() ?? string.Empty;
                    string lowerKvNum = lower.LookupParameter("КВ_Номер")?.AsString() ?? string.Empty;

                    bool isLivingLower = WetZoneCategories.LivingRooms.Contains(lowerType);
                    bool isKitchenLower = WetZoneCategories.KitchenRooms.Contains(lowerType);
                    if (!isLivingLower && !isKitchenLower) continue;

                    var lowerPoly = GetRoom2DOutline(lower);
                    List<Room> overlapping = upperRooms
                        .Where(upper =>
                        {
                            var upperPoly = GetRoom2DOutline(upper);
                            return DoPolygonsIntersect(lowerPoly, upperPoly);
                        })
                        .ToList();

                    foreach (var upper in overlapping)
                    {
                        string upperType = upper.LookupParameter(_selectedParam)?.AsString() ?? string.Empty;
                        string upperKvNum = upper.LookupParameter("КВ_Номер")?.AsString() ?? string.Empty;

                        bool isWetUpper = WetZoneCategories.WetRooms.Contains(upperType);
                        bool isKitchenUpper = WetZoneCategories.KitchenRooms.Contains(upperType);

                        // НЕЛЬЗЯ: Кухни над жилыми
                        if (isLivingLower && isKitchenUpper)
                            KitchenOverLiving_Illegal.Add(new List<Element> { lower, upper });

                        // НЕЛЬЗЯ: Мокрые над жилыми
                        if (isLivingLower && isWetUpper)
                            WetOverLiving_Illegal.Add(new List<Element> { lower, upper });

                        // МОЖНО (одна квартира) /НЕЛЬЗЯ (разные): Мокрые над кухнями
                        if (isKitchenLower && isWetUpper)
                        {
                            if (lowerKvNum == upperKvNum)
                                KitchenUnderWet_Accepted.Add(new List<Element> { lower, upper });
                            else
                                KitchenUnderWet_Illegal.Add(new List<Element> { lower, upper });
                        }
                    }
                }
            }

            // НЕЛЬЗЯ: InvalidEquipment над жилыми
            foreach (var upperFloor in familyInstancesByLevel.Keys)
            {
                foreach (var familyInstance in familyInstancesByLevel[upperFloor])
                {
                    int lowerFloor = upperFloor - 1;
                    if (!roomsByFloorParam.ContainsKey(lowerFloor))
                        continue;

                    List<Room> lowerRooms = roomsByFloorParam[lowerFloor].OfType<Room>().ToList();

                    foreach (var lower in lowerRooms)
                    {
                        string lowerType = lower.LookupParameter(_selectedParam)?.AsString() ?? string.Empty;
                        bool isLivingLower = WetZoneCategories.LivingRooms.Contains(lowerType);
                        if (!isLivingLower)
                            continue;

                        var lowerPoly = GetRoom2DOutline(lower);
                        var familyInstancePoly = GetFamilyInstance2DOutline(familyInstance);

                        if (DoPolygonsIntersect(familyInstancePoly, lowerPoly))
                        {
                            InvalidEquipmentOverLiving_Illegal.Add(new List<Element> { lower, familyInstance });
                        }
                    }
                }
            }












            var windowResult = new WetZoneResult(_uiDoc, roomsByFloorParam, KitchenOverLiving_Illegal, WetOverLiving_Illegal, KitchenUnderWet_Illegal, InvalidEquipmentOverLiving_Illegal, _selectedParam);
            windowResult.Show();            
        }











        /////////////////// Расчёт пересечения
        ///////////////////
        // Вспомогательный метод. Получение 2D-контуров. Помещение
        private static List<XYZ> GetRoom2DOutline(Room room)
        {
            var options = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            };

            var boundaries = room.GetBoundarySegments(options);
            if (boundaries == null || boundaries.Count == 0)
                return null;

            var outline = new List<XYZ>();
            var plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero);

            foreach (var seg in boundaries[0])
            {
                var pt = seg.GetCurve().GetEndPoint(0);

                double z = plane.Normal.DotProduct(pt - plane.Origin); 
                XYZ projected = pt - z * plane.Normal; 

                outline.Add(projected);
            }

            return outline;
        }

        // Вспомогательный метод. Получение 2D-контуров. Семейство
        private static List<XYZ> GetFamilyInstance2DOutline(FamilyInstance familyInstance)
        {
            var boundingBox = familyInstance.get_BoundingBox(null);
            if (boundingBox == null)
                return new List<XYZ>();

            var outline = new List<XYZ>
            {
                new XYZ(boundingBox.Min.X, boundingBox.Min.Y, 0),
                new XYZ(boundingBox.Max.X, boundingBox.Min.Y, 0),
                new XYZ(boundingBox.Max.X, boundingBox.Max.Y, 0),
                new XYZ(boundingBox.Min.X, boundingBox.Max.Y, 0)
            };

            return outline;
        }

        // Вспомогательный метод. Функция пересечения 2D-многоугольников
        private static bool DoPolygonsIntersect(List<XYZ> poly1, List<XYZ> poly2)
        {
            if (poly1 == null || poly2 == null || poly1.Count < 2 || poly2.Count < 2)
                return false;

            for (int i = 0; i < poly1.Count; i++)
            {
                XYZ a1 = poly1[i];
                XYZ a2 = poly1[(i + 1) % poly1.Count];

                for (int j = 0; j < poly2.Count; j++)
                {
                    XYZ b1 = poly2[j];
                    XYZ b2 = poly2[(j + 1) % poly2.Count];

                    if (LinesIntersect(a1, a2, b1, b2))
                        return true;
                }
            }

            if (IsPointInsidePolygon(poly1[0], poly2) || IsPointInsidePolygon(poly2[0], poly1))
                return true;

            return false;
        }

        // Вспомогательный метод. Проверка пересечения отрезков
        private static bool LinesIntersect(XYZ p1, XYZ p2, XYZ q1, XYZ q2)
        {
            double o1 = Orientation(p1, p2, q1);
            double o2 = Orientation(p1, p2, q2);
            double o3 = Orientation(q1, q2, p1);
            double o4 = Orientation(q1, q2, p2);

            return o1 * o2 < 0 && o3 * o4 < 0;
        }

        // Вспомогательный метод. Векторное произведение: (b - a) × (c - a)
        private static double Orientation(XYZ a, XYZ b, XYZ c)
        {
            return (b.X - a.X) * (c.Y - a.Y) - (b.Y - a.Y) * (c.X - a.X);
        }

        private static bool IsPointInsidePolygon(XYZ point, List<XYZ> polygon)
        {
            bool inside = false;
            int count = polygon.Count;

            for (int i = 0, j = count - 1; i < count; j = i++)
            {
                if (((polygon[i].Y > point.Y) != (polygon[j].Y > point.Y)) &&
                    (point.X < (polygon[j].X - polygon[i].X) * (point.Y - polygon[i].Y) /
                     (polygon[j].Y - polygon[i].Y) + polygon[i].X))
                {
                    inside = !inside;
                }
            }

            return inside;
        }
        /////////////////// 
        ///////////////////
    }
}
