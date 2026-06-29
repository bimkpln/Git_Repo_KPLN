using KPLN_CommandsWheel.ExternalCommands;
using KPLN_CommandsWheel.Models;
using KPLN_CommandsWheel.Services;
using KPLN_Library_PluginActivityWorker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace KPLN_CommandsWheel.Forms
{
    internal class CommandsWheelWindow : Window
    {
        private static CommandsWheelWindow _current;

        private readonly List<RevitCommandInfo> _commands;
        private readonly RevitCommandExecutor _executor;
        private bool _canCloseOnDeactivate;
        private bool _isClosing;

        internal CommandsWheelWindow(IEnumerable<RevitCommandInfo> commands, RevitCommandExecutor executor)
        {
            _commands = commands.Take(8).ToList();
            _executor = executor;
            _current = this;

            Title = "\u0428\u0442\u0443\u0440\u0432\u0430\u043b \u043a\u043e\u043c\u0430\u043d\u0434";
            Width = 260;
            Height = 260;
            ResizeMode = ResizeMode.NoResize;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            ShowInTaskbar = false;
            Topmost = true;
            Focusable = true;

            Closing += delegate
            {
                _isClosing = true;
                _canCloseOnDeactivate = false;
            };
            Closed += delegate
            {
                if (ReferenceEquals(_current, this))
                {
                    _current = null;
                }
            };
            Loaded += delegate
            {
                Activate();
                Focus();
                Keyboard.Focus(this);
                Dispatcher.BeginInvoke(new Action(delegate
                {
                    _canCloseOnDeactivate = true;
                }), DispatcherPriority.ApplicationIdle);
            };
            Deactivated += delegate
            {
                if (_canCloseOnDeactivate && !_isClosing)
                    CloseWheel();
            };
            KeyDown += delegate (object sender, KeyEventArgs args)
            {
                if (args.Key == Key.Escape)
                {
                    args.Handled = true;
                    CloseWheel();
                }
            };

            Content = BuildContent();
        }

        internal static bool TryActivateExisting()
        {
            if (_current == null || !_current.IsVisible)
                return false;

            if (_current.WindowState == WindowState.Minimized)
                _current.WindowState = WindowState.Normal;

            _current.Activate();
            return true;
        }

        private UIElement BuildContent()
        {
            Grid root = new Grid { Background = Brushes.Transparent };
            root.MouseLeftButtonDown += Root_MouseLeftButtonDown;

            Canvas canvas = new Canvas
            {
                Width = 260,
                Height = 260,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            double center = 130;
            double outerRadius = 108;
            double innerRadius = 39;
            double iconRadius = 73;
            double plateRadius = outerRadius;

            Ellipse plate = new Ellipse
            {
                Width = plateRadius * 2,
                Height = plateRadius * 2,
                Fill = new SolidColorBrush(Color.FromArgb(138, 24, 28, 34)),
                Stroke = new SolidColorBrush(Color.FromArgb(175, 210, 220, 230)),
                StrokeThickness = 1,
                Cursor = Cursors.SizeAll,
                ToolTip = "\u041f\u0435\u0440\u0435\u0442\u0430\u0449\u0438\u0442\u044c"
            };
            plate.MouseLeftButtonDown += delegate (object sender, MouseButtonEventArgs args)
            {
                StartDrag(args);
            };
            Canvas.SetLeft(plate, center - plateRadius);
            Canvas.SetTop(plate, center - plateRadius);
            canvas.Children.Add(plate);

            if (_commands.Count > 0)
            {
                List<UIElement> icons = new List<UIElement>();
                double segmentAngle = 360.0 / _commands.Count;
                double gapAngle = _commands.Count > 1 ? 1.5 : 0;

                for (int index = 0; index < _commands.Count; index++)
                {
                    RevitCommandInfo command = _commands[index];
                    double startAngle = -90 + index * segmentAngle + gapAngle / 2;
                    double endAngle = -90 + (index + 1) * segmentAngle - gapAngle / 2;

                    Path slice = CreateCommandSlice(command, center, innerRadius, outerRadius, startAngle, endAngle);
                    canvas.Children.Add(slice);

                    UIElement icon = CreateCommandIcon(command);
                    Point iconPoint = GetCirclePoint(center, iconRadius, (startAngle + endAngle) / 2);
                    Canvas.SetLeft(icon, iconPoint.X - 18);
                    Canvas.SetTop(icon, iconPoint.Y - 18);
                    icons.Add(icon);
                }

                foreach (UIElement icon in icons)
                {
                    canvas.Children.Add(icon);
                }
            }

            Border closeButton = CreateCloseButton(innerRadius * 2);
            Canvas.SetLeft(closeButton, center - closeButton.Width / 2);
            Canvas.SetTop(closeButton, center - closeButton.Height / 2);
            canvas.Children.Add(closeButton);

            root.Children.Add(canvas);
            return root;
        }

        private Path CreateCommandSlice(
            RevitCommandInfo command,
            double center,
            double innerRadius,
            double outerRadius,
            double startAngle,
            double endAngle)
        {
            Path slice = new Path
            {
                Data = CreateSliceGeometry(center, innerRadius, outerRadius, startAngle, endAngle),
                Fill = command.CanPost
                    ? new SolidColorBrush(Color.FromArgb(224, 38, 111, 166))
                    : new SolidColorBrush(Color.FromArgb(224, 88, 88, 88)),
                Stroke = new SolidColorBrush(Color.FromArgb(150, 230, 235, 240)),
                StrokeThickness = 1,
                ToolTip = command.Name,
                Cursor = Cursors.Hand
            };

            slice.MouseLeftButtonDown += delegate (object sender, MouseButtonEventArgs args)
            {
                args.Handled = true;
            };
            slice.MouseEnter += delegate
            {
                ApplySliceState(slice, command, true);
            };
            slice.MouseLeave += delegate
            {
                ApplySliceState(slice, command, false);
            };
            slice.MouseLeftButtonUp += delegate (object sender, MouseButtonEventArgs args)
            {
                args.Handled = true;
                RunAndClose(command);
            };

            return slice;
        }

        private UIElement CreateCommandIcon(RevitCommandInfo command)
        {
            Border host = new Border
            {
                Width = 36,
                Height = 36,
                IsHitTestVisible = false
            };

            ImageSource icon = IconSourceLoader.Load(command);
            if (icon != null)
            {
                host.Child = new Image
                {
                    Source = icon,
                    Width = 28,
                    Height = 28,
                    Stretch = Stretch.Uniform,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };

                return host;
            }

            host.Child = new TextBlock
            {
                Text = string.IsNullOrWhiteSpace(command.Name) ? "?" : command.Name.Substring(0, 1).ToUpperInvariant(),
                Foreground = Brushes.White,
                FontSize = 21,
                FontWeight = FontWeights.SemiBold,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Center
            };

            return host;
        }

        private Border CreateCloseButton(double size)
        {
            Border button = new Border
            {
                Width = size,
                Height = size,
                CornerRadius = new CornerRadius(size / 2),
                Background = new SolidColorBrush(Color.FromRgb(190, 45, 58)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(238, 160, 168)),
                BorderThickness = new Thickness(1),
                ToolTip = "\u0417\u0430\u043a\u0440\u044b\u0442\u044c",
                Cursor = Cursors.Hand,
                Child = new TextBlock
                {
                    Text = "\u00D7",
                    FontSize = 30,
                    Foreground = Brushes.White,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, -4, 0, 0)
                }
            };

            button.MouseLeftButtonDown += delegate (object sender, MouseButtonEventArgs args)
            {
                args.Handled = true;
            };
            button.MouseEnter += delegate
            {
                button.Background = new SolidColorBrush(Color.FromRgb(215, 58, 72));
            };
            button.MouseLeave += delegate
            {
                button.Background = new SolidColorBrush(Color.FromRgb(190, 45, 58));
            };
            button.MouseLeftButtonUp += delegate (object sender, MouseButtonEventArgs args)
            {
                args.Handled = true;
                CloseWheel();
            };

            return button;
        }

        private void ApplySliceState(Path slice, RevitCommandInfo command, bool isHover)
        {
            if (isHover)
            {
                slice.Fill = command.CanPost
                    ? new SolidColorBrush(Color.FromArgb(246, 62, 144, 207))
                    : new SolidColorBrush(Color.FromArgb(246, 110, 110, 110));
                slice.Stroke = Brushes.White;
                slice.StrokeThickness = 2.2;
                return;
            }

            slice.Fill = command.CanPost
                ? new SolidColorBrush(Color.FromArgb(224, 38, 111, 166))
                : new SolidColorBrush(Color.FromArgb(224, 88, 88, 88));
            slice.Stroke = new SolidColorBrush(Color.FromArgb(150, 230, 235, 240));
            slice.StrokeThickness = 1;
        }

        private void RunAndClose(RevitCommandInfo command)
        {
            _executor.Run(command);
            
            DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(CommandsWheel.PluginName, ModuleData.ModuleName).ConfigureAwait(false);

            CloseWheel();
        }

        private void CloseWheel()
        {
            if (_isClosing)
                return;

            _isClosing = true;
            _canCloseOnDeactivate = false;
            Close();
        }

        private static Geometry CreateSliceGeometry(double center, double innerRadius, double outerRadius, double startAngle, double endAngle)
        {
            if (endAngle - startAngle >= 359.9)
            {
                endAngle = startAngle + 359.5;
            }

            Point outerStart = GetCirclePoint(center, outerRadius, startAngle);
            Point outerEnd = GetCirclePoint(center, outerRadius, endAngle);
            Point innerEnd = GetCirclePoint(center, innerRadius, endAngle);
            Point innerStart = GetCirclePoint(center, innerRadius, startAngle);
            bool isLargeArc = endAngle - startAngle > 180;

            PathFigure figure = new PathFigure
            {
                StartPoint = outerStart,
                IsClosed = true
            };
            figure.Segments.Add(new ArcSegment(outerEnd, new Size(outerRadius, outerRadius), 0, isLargeArc, SweepDirection.Clockwise, true));
            figure.Segments.Add(new LineSegment(innerEnd, true));
            figure.Segments.Add(new ArcSegment(innerStart, new Size(innerRadius, innerRadius), 0, isLargeArc, SweepDirection.Counterclockwise, true));

            PathGeometry geometry = new PathGeometry();
            geometry.Figures.Add(figure);
            return geometry;
        }

        private static Point GetCirclePoint(double center, double radius, double angleDegrees)
        {
            double angle = Math.PI * angleDegrees / 180.0;
            return new Point(
                center + Math.Cos(angle) * radius,
                center + Math.Sin(angle) * radius);
        }

        private void StartDrag(MouseButtonEventArgs args)
        {
            if (args.ChangedButton != MouseButton.Left)
                return;

            args.Handled = true;
            try
            {
                DragMove();
            }
            catch
            {
                return;
            }
        }

        private void Root_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            StartDrag(e);
        }
    }
}
