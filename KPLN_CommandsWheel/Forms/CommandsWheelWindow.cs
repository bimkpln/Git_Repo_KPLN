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
        private readonly bool _isPinned;
        private readonly bool _showCloseButton;
        private bool _canCloseOnDeactivate;
        private bool _isClosing;
        private bool _isCommandPressActive;
        private bool _isWindowDragActive;
        private Point _commandPressPoint;
        private RevitCommandInfo _pressedCommand;
        private UIElement _pressedElement;
        private int _selectedIndex = -1;

        internal CommandsWheelWindow(IEnumerable<RevitCommandInfo> commands, RevitCommandExecutor executor, UserSettings settings)
        {
            _commands = commands.Take(8).ToList();
            _executor = executor;
            _isPinned = settings != null && string.Equals(settings.WheelMode, WheelModeNames.Pinned, StringComparison.OrdinalIgnoreCase);
            _showCloseButton = _isPinned || (settings != null && settings.IsWheelCloseButtonVisible);
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
                if (!_isPinned && _canCloseOnDeactivate && !_isClosing)
                {
                    CloseWheel();
                }
            };
            KeyDown += delegate (object sender, KeyEventArgs args)
            {
                if (args.Key == Key.Escape)
                {
                    args.Handled = true;
                    if (!_isPinned)
                    {
                        CloseWheel();
                    }

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
                Fill = new SolidColorBrush(Color.FromArgb(18, 140, 140, 140)),
                Stroke = new SolidColorBrush(Color.FromArgb(32, 180, 180, 180)),
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

            if (_showCloseButton)
            {
                Border closeButton = CreateCloseButton(innerRadius * 2);
                Canvas.SetLeft(closeButton, center - closeButton.Width / 2);
                Canvas.SetTop(closeButton, center - closeButton.Height / 2);
                canvas.Children.Add(closeButton);
            }

            Border dragHandle = CreateDragHandle();
            Canvas.SetLeft(dragHandle, center - dragHandle.Width / 2);
            Canvas.SetTop(dragHandle, 7);
            canvas.Children.Add(dragHandle);

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
                Fill = new SolidColorBrush(Color.FromArgb(42, 145, 145, 145)),
                Stroke = new SolidColorBrush(Color.FromArgb(88, 235, 235, 235)),
                StrokeThickness = 1,
                ToolTip = command.Name,
                Cursor = Cursors.Hand
            };

            slice.MouseLeftButtonDown += delegate (object sender, MouseButtonEventArgs args)
            {
                args.Handled = true;
                if (_isPinned)
                {
                    BeginCommandPress(command, slice, args);
                }
            };
            slice.MouseMove += delegate (object sender, MouseEventArgs args)
            {
                if (_isPinned)
                {
                    TrackCommandDrag(slice, args);
                }
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
                if (_isPinned)
                {
                    CompleteCommandPress(command);
                    return;
                }

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
                ToolTip = "Закрыть",
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

        private Border CreateDragHandle()
        {
            Grid grip = new Grid();
            grip.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grip.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grip.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            for (int index = 0; index < 3; index++)
            {
                Border line = new Border
                {
                    Height = 1,
                    Margin = new Thickness(7, 0, 7, 0),
                    Background = new SolidColorBrush(Color.FromArgb(135, 245, 245, 245)),
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetRow(line, index);
                grip.Children.Add(line);
            }

            Border handle = new Border
            {
                Width = 40,
                Height = 14,
                CornerRadius = new CornerRadius(4),
                Background = new SolidColorBrush(Color.FromArgb(175, 42, 42, 42)),
                BorderBrush = new SolidColorBrush(Color.FromArgb(90, 255, 255, 255)),
                BorderThickness = new Thickness(1),
                Cursor = Cursors.SizeAll,
                ToolTip = "\u041f\u0435\u0440\u0435\u0442\u0430\u0449\u0438\u0442\u044c",
                Child = grip
            };

            handle.MouseLeftButtonDown += delegate (object sender, MouseButtonEventArgs args)
            {
                StartDrag(args);
            };

            return handle;
        }

        private void ApplySliceState(Path slice, RevitCommandInfo command, bool isSelected)
        {
            if (isSelected)
            {
                slice.Fill = new SolidColorBrush(Color.FromArgb(104, 118, 118, 118));
                slice.Stroke = Brushes.White;
                slice.StrokeThickness = 2.2;
                return;
            }

            slice.Fill = new SolidColorBrush(Color.FromArgb(42, 145, 145, 145));
            slice.Stroke = new SolidColorBrush(Color.FromArgb(88, 235, 235, 235));
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

        private void BeginCommandPress(RevitCommandInfo command, UIElement element, MouseButtonEventArgs args)
        {
            if (args.ChangedButton != MouseButton.Left)
            {
                return;
            }

            _pressedCommand = command;
            _pressedElement = element;
            _commandPressPoint = args.GetPosition(this);
            _isCommandPressActive = true;
            _isWindowDragActive = false;
            element.CaptureMouse();
        }

        private void TrackCommandDrag(UIElement element, MouseEventArgs args)
        {
            if (!_isCommandPressActive || args.LeftButton != MouseButtonState.Pressed)
            {
                return;
            }

            Point currentPoint = args.GetPosition(this);
            if (Math.Abs(currentPoint.X - _commandPressPoint.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(currentPoint.Y - _commandPressPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            _isWindowDragActive = true;
            _isCommandPressActive = false;
            element.ReleaseMouseCapture();
            DragWindow();
        }

        private void CompleteCommandPress(RevitCommandInfo command)
        {
            if (_pressedElement != null)
            {
                _pressedElement.ReleaseMouseCapture();
            }

            bool shouldRun = _isCommandPressActive
                && !_isWindowDragActive
                && ReferenceEquals(_pressedCommand, command);

            _pressedCommand = null;
            _pressedElement = null;
            _isCommandPressActive = false;
            _isWindowDragActive = false;

            if (shouldRun)
            {
                RunAndClose(command);
            }
        }

        private void RunAndClose(RevitCommandInfo command)
        {
            _executor.Run(command);

            if (!_isPinned)
            {
                CloseWheel();
            }
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
            DragWindow();
        }

        private void DragWindow()
        {
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