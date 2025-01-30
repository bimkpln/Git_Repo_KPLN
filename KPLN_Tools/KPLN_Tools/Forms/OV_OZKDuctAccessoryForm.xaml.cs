using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.ExecutableCommand;
using KPLN_Tools.Common.OVVK_System;
using KPLN_Tools.ExecutableCommand;
using System;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class OV_OZKDuctAccessoryForm : Window
    {
        private readonly UIApplication _uiapp;

        public OV_OZKDuctAccessoryForm(
            UIApplication uiapp,
            OZKDuctAccessoryEntity[] ozkDuctAccessoryEntities)
        {
            _uiapp = uiapp;
            OZKDuctAccessoryEntities = ozkDuctAccessoryEntities;

            InitializeComponent();
            OZKTypes.ItemsSource = OZKDuctAccessoryEntities;

            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        public OZKDuctAccessoryEntity[] OZKDuctAccessoryEntities { get; private set; }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
            }
        }

        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandOZKDuctAccessory_Start(OZKDuctAccessoryEntities));
            Close();
        }

        private void RefreshBtn_Click(object sender, RoutedEventArgs e)
        {
            // Создаем тип
            Type type = Type.GetType(typeof(ExternalCommands.Command_OV_OZKDuctAccessory).FullName, true);

            // Создаем экземпляр типа
            object instance = Activator.CreateInstance(type);

            // Определяем метод ExecuteByUIApp
            MethodInfo executeMethod = type.GetMethod("ExecuteByUIApp");

            // Вызываем метод ExecuteByUIApp, передавая _uiApp как аргумент
            if (executeMethod != null)
                executeMethod.Invoke(instance, new object[] { _uiapp });
            else
                throw new Exception("Ошибка определения метода через рефлексию. Отправь это разработчику\n");

            Close();
        }

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Выполните здесь нужное вам поведение при прокрутке колесом мыши
            ScrollViewer.ScrollToVerticalOffset(ScrollViewer.VerticalOffset - e.Delta);
            // Пометьте событие как обработанное, чтобы оно не передалось другим элементам
            e.Handled = true;
        }

        private void ItemBtn_Click(object sender, RoutedEventArgs e)
        {
            Button btn = (Button)sender;
            OZKDuctAccessoryEntity currentEntity = btn.DataContext as OZKDuctAccessoryEntity;
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ZoomElementCommand(currentEntity.CurrentFamilyInstances));
        }
    }
}
