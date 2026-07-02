using KPLN_CommandsWheel.Models;
using KPLN_CommandsWheel.Services;
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
        private readonly List<Path> _slices = new List<Path>();
        private bool _canCloseOnDeactivate;
        private bool _isClosing;
        private int _selectedIndex = -1;

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
                {
                    CloseWheel();
                }
            };
            KeyDown += delegate (object sender, KeyEventArgs args)
            {
                if (args.Key == Key.Escape)
                {
                    args.Handled = true;
                    CloseWheel();
                    return;
                }

                if (args.Key == Key.Left)
                {
                    args.Handled = true;
                    MoveSelection(-1);
                    return;
                }

                if (args.Key == Key.Right)
                {
                    args.Handled = true;
                    MoveSelection(1);
                    return;
                }

                if (args.Key == Key.Enter || args.Key == Key.Return || args.Key == Key.Space)
                {
                    args.Handled = true;
                    RunSelectedCommand();
                }
            };

            Content = BuildContent();
        }

        internal static bool TryActivateExisting()
        {
            if (_current == null || !_current.IsVisible)
            {
                return false;
            }

            if (_current.WindowState == WindowState.Minimized)
            {
                _current.WindowState = WindowState.Normal;
            }

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
                Fill = new SolidColorBrush(Color.FromArgb(16, 24, 28, 34)),
                Stroke = new SolidColorBrush(Color.FromArgb(38, 210, 220, 230)),
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
                _slices.Clear();
                _selectedIndex = 0;

                List<UIElement> icons = new List<UIElement>();
                double segmentAngle = 360.0 / _commands.Count;
                double gapAngle = _commands.Count > 1 ? 1.5 : 0;

                for (int index = 0; index < _commands.Count; index++)
                {
                    RevitCommandInfo command = _commands[index];
                    double startAngle = -90 + index * segmentAngle + gapAngle / 2;
                    double endAngle = -90 + (index + 1) * segmentAngle - gapAngle / 2;

                    Path slice = CreateCommandSlice(command, index, center, innerRadius, outerRadius, startAngle, endAngle);
                    _slices.Add(slice);
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

                RefreshSliceStates();
            }

            root.Children.Add(canvas);
            return root;
        }

        private Path CreateCommandSlice(
            RevitCommandInfo command,
            int index,
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
                _selectedIndex = index;
                RefreshSliceStates();
            };
            slice.MouseLeave += delegate
            {
                RefreshSliceStates();
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

        private void ApplySliceState(Path slice, RevitCommandInfo command, bool isSelected)
        {
            if (isSelected)
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

        private void RefreshSliceStates()
        {
            for (int index = 0; index < _slices.Count && index < _commands.Count; index++)
            {
                ApplySliceState(_slices[index], _commands[index], index == _selectedIndex);
            }
        }

        private void MoveSelection(int direction)
        {
            if (_commands.Count == 0)
            {
                return;
            }

            if (_selectedIndex < 0 || _selectedIndex >= _commands.Count)
            {
                _selectedIndex = 0;
            }
            else
            {
                _selectedIndex = (_selectedIndex + direction + _commands.Count) % _commands.Count;
            }

            RefreshSliceStates();
        }

        private void RunSelectedCommand()
        {
            if (_selectedIndex < 0 || _selectedIndex >= _commands.Count)
            {
                return;
            }

            RunAndClose(_commands[_selectedIndex]);
        }

        private void RunAndClose(RevitCommandInfo command)
        {
            _executor.Run(command);
            CloseWheel();
        }

        private void CloseWheel()
        {
            if (_isClosing)
            {
                return;
            }

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
            {
                return;
            }

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
