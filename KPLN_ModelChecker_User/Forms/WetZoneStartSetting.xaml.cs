using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Логика взаимодействия для WetZoneStartSetting.xaml
    /// </summary>
    public partial class WetZoneReviewWindow : Window
    {
        public WetZoneReviewWindow(Document doc, List<Element> livingRooms, List<Element> kitchenRooms, List<Element> wetRooms, List<Element> undefinedRooms)
        {
            InitializeComponent();

            LivingExp.Header = $"Жилые комнаты ({livingRooms.Count})";
            WetExp.Header = $"Мокрые зоны ({wetRooms.Count})";
            KitchenExp.Header = $"Кухни ({kitchenRooms.Count})";
            UndefinedExp.Header = $"Неопределённые помещения ({undefinedRooms.Count})";

            LivingList.ItemsSource = FormatRooms(livingRooms);
            WetList.ItemsSource = FormatRooms(wetRooms);
            KitchenList.ItemsSource = FormatRooms(kitchenRooms);           
            UndefinedList.ItemsSource = FormatRooms(undefinedRooms);

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
            bool hasNotCriticallError = false;


            // Проверка помещений
            List<Element> allRooms = livingRooms.Concat(kitchenRooms).Concat(wetRooms).ToList();
            if (allRooms.Count == 0)
            {
                report.Inlines.Add(new Run("❎ Нет доступных элементов в категориях 'Жилые комнаты', 'Мокрые зоны', 'Кухни'.\n")
                {
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.Red
                });
                //hasCriticalError = true;
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
                        results[levelName] = "ОШИБКА";
                        //hasCriticalError = true;
                    }
                }

                List<Inline> output = new List<Inline>();

                output.Add(new Run("🔎 Результаты преобразования названия уровней в этажи:\n") { FontWeight = FontWeights.SemiBold });

                foreach (var pair in results)
                {
                    bool isError = pair.Value == "ОШИБКА";
                    output.Add(new Run($"       • {pair.Key} → {pair.Value}\n")
                    {
                        Foreground = isError ? Brushes.Red : Brushes.DarkGreen
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
                    Foreground = Brushes.BlueViolet
                });

                foreach (var pair in invalidLevelElements)
                {
                    string display = $"       • {pair.Value} [ID {pair.Key.Id.IntegerValue}]\n";
                    outputInvalidLevelElements.Add(new Run(display)
                    {
                        Foreground = Brushes.BlueViolet
                    });
                }

                foreach (var inline in outputInvalidLevelElements)
                {
                    report.Inlines.Add(inline);
                }
                
                //hasOtherlError = true;
            }

            report.Inlines.Add(new LineBreak());

            // Заполненость параметра КВ_Номер
            List<Element> missingKv = allRooms.Where(el =>
            {
                var param = el.LookupParameter("КВ_Номер");
                return param == null || string.IsNullOrWhiteSpace(param.AsString());
            }).ToList();

            if (missingKv.Count > 0)
            {
                report.Inlines.Add(new Run("⚠️ Внимание: у следующих помещений не заполнен параметр 'КВ_Номер':\n")
                {
                    FontWeight = FontWeights.SemiBold,
                    Foreground = Brushes.Red
                });

                foreach (Element el in missingKv)
                {
                    string name = el.Name ?? "<без имени>";
                    string id = el.Id.IntegerValue.ToString();

                    report.Inlines.Add(new Run($"       • {name} [ID {id}]\n")
                    {
                        Foreground = Brushes.Red
                    });
                }

                //hasOtherlError = true;
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





            //if (hasCriticalError)
            //{
            //   report.Inlines.Add(new Run("⛔ Плагин не может продолжить работу.")
            //    {
            //        Foreground = Brushes.Red,
            //        FontWeight = FontWeights.Bold
            //    });
           //     Start1Button.IsEnabled = false;
           // }
           // else if (hasOtherlError)
           // {
           //     report.Inlines.Add(new Run("⚠ Есть ошибки, но запуск плагина возможен.")
           //     {
          //          Foreground = Brushes.DarkOrange,
           //         FontWeight = FontWeights.Bold
           //     });
          //      Start1Button.IsEnabled = true;
         //   }
         //   else
         //   {
          //      report.Inlines.Add(new Run("✅ Все проверки пройдены. Готово к запуску.")
          //      {
          //          Foreground = Brushes.DarkGreen,
        //            FontWeight = FontWeights.Bold
          //      });
        //        Start1Button.IsEnabled = true;
        //    }
        //
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

        // XAML. Отмена
        private void ButtonCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
