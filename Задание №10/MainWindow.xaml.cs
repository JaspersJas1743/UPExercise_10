using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.Collections.Generic;
using Excel = NetOffice.ExcelApi;

namespace Задание__10
{
	public partial class MainWindow : Window
	{
		private TextBox _lastFocused;
		private Excel.Application _excel = new Excel.Application();
		private Dictionary<string, int> _limitForBase = new Dictionary<string, int>()
		{
			{ "Bin", 2 },
			{ "Oct", 8 },
			{ "Dec", 10 },
			{ "Hex", 16 }
		};

		public MainWindow()
		{
			InitializeComponent();
			Bin.Focus();
		}

		private void BinTextChanged(object sender, TextChangedEventArgs e)
		{
			TextBox tb = (TextBox)sender;
			try
			{
				RecoveryIfChangeIsRemove(e, tb);
				Oct.Text = _excel.WorksheetFunction.Bin2Oct(Bin.Text);
				Dec.Text = _excel.WorksheetFunction.Bin2Dec(Bin.Text);
				Hex.Text = _excel.WorksheetFunction.Bin2Hex(Bin.Text);
			} catch
			{
				ExcelError(tb);
			}
		}

		private void OctTextChanged(object sender, TextChangedEventArgs e)
		{
			TextBox tb = (TextBox)sender;
			try
			{
				RecoveryIfChangeIsRemove(e, tb);
				Bin.Text = _excel.WorksheetFunction.Oct2Bin(Oct.Text);
				Dec.Text = _excel.WorksheetFunction.Oct2Dec(Oct.Text);
				Hex.Text = _excel.WorksheetFunction.Oct2Hex(Oct.Text);
			} catch
			{
				ExcelError(tb);
			}
		}

		private void DecTextChanged(object sender, TextChangedEventArgs e)
		{
			TextBox tb = (TextBox)sender;
			try
			{
				RecoveryIfChangeIsRemove(e, tb);
				Bin.Text = _excel.WorksheetFunction.Dec2Bin(Dec.Text);
				Oct.Text = _excel.WorksheetFunction.Dec2Oct(Dec.Text);
				Hex.Text = _excel.WorksheetFunction.Dec2Hex(Dec.Text);
			} catch
			{
				ExcelError(tb);
			}
		}

		private void HexTextChanged(object sender, TextChangedEventArgs e)
		{
			TextBox tb = (TextBox)sender;
			try
			{
				RecoveryIfChangeIsRemove(e, tb);
				Bin.Text = _excel.WorksheetFunction.Hex2Bin(Hex.Text);
				Oct.Text = _excel.WorksheetFunction.Hex2Oct(Hex.Text);
				Dec.Text = _excel.WorksheetFunction.Hex2Dec(Hex.Text);
			} catch
			{
				ExcelError(tb);
			}
		}

		private void TextBoxPreviewTextInput(object sender, TextCompositionEventArgs e)
			=> e.Handled = !Validate(((TextBox)sender).Name, e.Text);

		private void TextBoxGotFocus(object sender, RoutedEventArgs e)
			=> _lastFocused = (TextBox)sender;

		private static void RecoveryIfChangeIsRemove(TextChangedEventArgs e, TextBox tb)
		{
			if(e.Changes.Last().RemovedLength > 0 && (tb.Text.Length == 0))
				tb.Text = "0";
			tb.SelectionStart = tb.Text.Length;
		}

		private void NumberClick(object sender, RoutedEventArgs e)
		{
			string added = ((Button)sender).Content.ToString();
			if(Validate(_lastFocused.Name, added))
				_lastFocused.Text += added;
			_lastFocused.Focus();
		}

		private void ExitClick(object sender, RoutedEventArgs e)
		{
			if(MessageBox.Show("Вы уверены, что хотите закрыть приложение?", "Подтверждение действия", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
				return;
			_excel.Quit();
			Application.Current.Shutdown();
		}

		private void DeactivateClick(object sender, RoutedEventArgs e)
			=> Application.Current.MainWindow.WindowState = WindowState.Minimized;

		private static void ExcelError(TextBox tb)
		{
			tb.Text = "0";
			MessageBox.Show("Ошибка в работе функций MS Excel", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
		}

		private void DragMove(object sender, MouseButtonEventArgs e)
		{
			if(Mouse.LeftButton == MouseButtonState.Pressed)
				Application.Current.MainWindow.DragMove();
		}

		private bool Validate(string currentBase, string numToCheck)
		{
			try
			{
				_ = Convert.ToInt32(numToCheck, fromBase: _limitForBase[currentBase]);
				return true;
			} catch
			{
				return false;
			}
		}
	}
}
