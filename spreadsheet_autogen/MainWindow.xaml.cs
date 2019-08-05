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
using System.IO;

namespace spreadsheet_autogen
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
          
            InitializeComponent();
            
            this.Row.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            this.Column.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            this.MaxValue.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            this.MinValue.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            
            ///<summary>
            ///Запрещает вставку символов
            ///</summary>>
            void textBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
            {
                if (!Char.IsDigit(e.Text, 0)) e.Handled = true;
            }



        }

        private void Row_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
        /// <summary>
        /// Обработчик нажатия кнопки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void onClick(object sender, RoutedEventArgs e)
        {
            GenerateSheet sheet = new GenerateSheet();
            var excelPackage = sheet.ImportPackage();
            sheet.CreateWorksheet(excelPackage);

            if (Choose.SelectedValue == number)
            {
                sheet.CellsRandomNumbers(Row.Text, Column.Text, MinValue.Text, MaxValue.Text);
            }

            if (Choose.SelectedValue == @char)
            {
                //Це костыль
                string Random = "Test";
                sheet.CellRandomString(Random, Row.Text, Column.Text, CharLength.Text);
            }

            if (Choose.SelectedValue == user)
            {
                sheet.CellUserValue(Row.Text, Column.Text, UserValue.Text);
            }
        }

        private void Choose_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Choose.SelectedValue == number)
            {
                MaxValueLabel.Visibility = Visibility.Visible;
                MinValueLabel.Visibility = Visibility.Visible;
                MaxValue.Visibility = Visibility.Visible;
                MinValue.Visibility = Visibility.Visible;
                UserValue.Visibility = Visibility.Hidden;
                CharLength.Visibility = Visibility.Hidden;
                CharLabel.Visibility = Visibility.Hidden;
            }
            else if (Choose.SelectedValue == @char)
            {
                MaxValueLabel.Visibility = Visibility.Hidden;
                MinValueLabel.Visibility = Visibility.Hidden;
                MaxValue.Visibility = Visibility.Hidden;
                MinValue.Visibility = Visibility.Hidden;
                UserValue.Visibility = Visibility.Hidden;
                CharLength.Visibility = Visibility.Visible;
                CharLabel.Visibility = Visibility.Visible;
            }
            else if (Choose.SelectedValue == user)
            {
                MaxValueLabel.Visibility = Visibility.Hidden;
                MinValueLabel.Visibility = Visibility.Hidden;
                MaxValue.Visibility = Visibility.Hidden;
                MinValue.Visibility = Visibility.Hidden;
                UserValue.Visibility = Visibility.Visible;
                CharLength.Visibility = Visibility.Hidden;
                CharLabel.Visibility = Visibility.Hidden;
            }
        }

        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {

        }

        private void TextBox_TextChanged_2(object sender, TextChangedEventArgs e)
        {

        }

        
    }
}
