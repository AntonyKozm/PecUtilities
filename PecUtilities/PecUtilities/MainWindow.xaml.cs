using System.Windows;
using System;
using System.Windows.Controls;

namespace PecUtilities
{
    public partial class MainWindow : Window
    {
        private readonly ToastViewModel _vm;

        public MainWindow()
        {
            InitializeComponent();

            _vm = new ToastViewModel();
            Unloaded += OnUnload;

            btConvert.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); 
        }

        private void OnUnload(object sender, RoutedEventArgs e)
        {
            _vm.OnUnloaded();
        }

        private void btConvert_Click(object sender, RoutedEventArgs e)
        {
            main.Child = new Converter(_vm);
        }

        private void btMagazine_Click(object sender, RoutedEventArgs e)
        {
            main.Child = new MagazineCreator(_vm);
        }

        private void btDeleteNumbers_Click(object sender, RoutedEventArgs e)
        {
            main.Child = new DeleteNumbers();
        }

        #region DragAndDrop
        private void Window_PreviewDragEnter(object sender, DragEventArgs e)
        {
            if (main.Child is Converter)
            {
                ((Converter)main.Child).SetDragMaskIndex(1);
            }
        }

        private void Window_PreviewDragLeave(object sender, DragEventArgs e)
        {
            if (main.Child is Converter)
            {
                ((Converter)main.Child).SetDragMaskIndex(-1);
            }
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (main.Child is Converter)
            {
                ((Converter)main.Child).SetDragMaskIndex(-1);
            }
        }
        #endregion

        #region ToastMessages
        void ShowMessage(Action<string> action, string text)
        {
            action(text);
        }

        public ToastViewModel GetViewModel()
        {
            return _vm;
        }
        #endregion
    }
}
