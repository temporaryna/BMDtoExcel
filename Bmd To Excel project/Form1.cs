using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Bmd_To_Excel_project
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		// окно excel
		private Excel.Application excelapp;

		// листы и рабочий лист excel
		private Excel.Sheets excelsheets;
		private Excel.Worksheet excelworksheet;

		// выделенные ячейки для работы excel
		private Excel.Range excelcells;

		// потоковый файл BMD
		Stream BmdFile;

		const int ITEM_MAX_LINE_SIZE = 84;	// максимальный размер строки item.bmd
		const int ITEM_MAX_IN_TYPE = 512;	// максимальное количество вещей в одном типе item.bmd
		const int ITEM_MAX_TYPES = 16;		// максимальное количество типов в item.bmd
		// максимальный размер файла item.bmd - 688128
		const int ITEM_TOTAL_LINE_SIZE = ITEM_MAX_IN_TYPE * ITEM_MAX_TYPES * ITEM_MAX_LINE_SIZE;

		// структура для файла item.bmd
		unsafe public struct ItemEx
		{
			public string ItemName;
			public fixed Int64 Numbers[54];
		};
		public ItemEx[,] Items = new ItemEx[16, 512];

		// загрузка потока файла BMD в структуру
		unsafe private bool LoadItemBmd(Stream File)
		{
			byte[] lpDstBuf = new byte[ITEM_TOTAL_LINE_SIZE];	// массив для файла после декриптования
			byte[] XorKeys = new byte[3] { 0xFC, 0xCF, 0xAB };	// ключи для декриптования

			// обнуление, надо ли ?
			//for (int i = 0; i < ITEM_TOTAL_LINE_SIZE; i++)
			//	lpDstBuf[i] = 0;

			// считывание потока в массив
			File.Read(lpDstBuf, 0, ITEM_TOTAL_LINE_SIZE);

			// декриптование
			for (int i = 0; i < ITEM_TOTAL_LINE_SIZE; i++)
				lpDstBuf[i] ^= XorKeys[i % 3];

			// цикл по всем типам
			for (int i = 0; i < ITEM_MAX_TYPES; i++)
			{
				// цикл по всем вещам в типе
				for (int j = 0; j < ITEM_MAX_IN_TYPE; j++)
				{
					int start = (i * ITEM_MAX_IN_TYPE + j) * ITEM_MAX_LINE_SIZE;	// позиция строки в массиве
					// считываем название вещи (первые 30 символов строки)
					Items[i, j].ItemName = System.Text.Encoding.Default.GetString(lpDstBuf, start, 30);
					start += 30;	// перемещаем позицию на 30

					// фиксируем переменную структуры в буфер для изменения
					fixed (Int64* buff2 = Items[i, j].Numbers)
					{
						int TempX = 0;	// позиция в строке массива

						for (int k = 0; k < MyItemColumns.Length; k++)
						{
							// цикл по ячейкам в структуре буфера
							Int64* buff = buff2 + k;		// переменная в строке буфера
							int TempPos = start + TempX;	// позиция в массиве item.bmd

							// проверяем тип хранящейся переменной и копируем из массива item.bmd в буфер
							if (MyItemColumns[k].TypeSize == 1)
							{
								if (MyItemColumns[k].Signed)
									*buff = (sbyte)lpDstBuf[TempPos];
								else
									*buff = (byte)lpDstBuf[TempPos];
							}
							else if (MyItemColumns[k].TypeSize == 2)
							{
								if (MyItemColumns[k].Signed)
									*buff = BitConverter.ToInt16(lpDstBuf, TempPos);
								else
									*buff = BitConverter.ToUInt16(lpDstBuf, TempPos);
							}
							else if (MyItemColumns[k].TypeSize == 4)
							{
								if (MyItemColumns[k].Signed)
									*buff = BitConverter.ToInt32(lpDstBuf, TempPos);
								else
									*buff = BitConverter.ToUInt32(lpDstBuf, TempPos);
							}
							// увеличиваем позицию в массиве item.bmd на размер переменной
							TempX += MyItemColumns[k].TypeSize;
						}
					}
				}
			}
			return true;
		}

		// сохранение структуры в item.bmd
		unsafe private void SaveItemBmd()
		{
			if (saveFileDialog1.ShowDialog() == DialogResult.OK)
			{
				// подгатавливаем буфер для структуры вещей
				byte[] lpDstBuf = new byte[ITEM_TOTAL_LINE_SIZE + 4];
				byte[] XorKeys = new byte[3] { 0xFC, 0xCF, 0xAB };	// ключи для криптования

				// цикл по всем типам
				for (int i = 0; i < ITEM_MAX_TYPES; i++)
				{
					// цикл по всем вещам в типе
					for (int j = 0; j < ITEM_MAX_IN_TYPE; j++)
					{
						int start = (i * ITEM_MAX_IN_TYPE + j) * ITEM_MAX_LINE_SIZE;	// позиция строки в массиве

						// запись названия вещи (первые 30 символов строки)
						Array.Copy(System.Text.Encoding.Default.GetBytes(Items[i, j].ItemName), 0, lpDstBuf, start, Items[i, j].ItemName.Length);

						start += 30;	// перемещаем позицию на 30

						// фиксируем переменную структуры в буфер для изменения
						fixed (Int64* buff2 = Items[i, j].Numbers)
						{
							int TempX = 0;	// позиция в строке массива

							for (int k = 0; k < MyItemColumns.Length; k++)
							{
								// цикл по ячейкам в структуре буфера
								Int64* buff = buff2 + k;		// переменная в строке буфера
								int TempPos = start + TempX;	// позиция в массиве item.bmd

								// проверяем тип хранящейся переменной и копируем из массива item.bmd в буфер
								if (MyItemColumns[k].TypeSize == 1)
								{
									if (MyItemColumns[k].Signed)
										Array.Copy(BitConverter.GetBytes(Convert.ToSByte(*buff)), 0, lpDstBuf, TempPos, 1);
									else
										Array.Copy(BitConverter.GetBytes(Convert.ToByte(*buff)), 0, lpDstBuf, TempPos, 1);
								}
								else if (MyItemColumns[k].TypeSize == 2)
								{
									if (MyItemColumns[k].Signed)
										Array.Copy(BitConverter.GetBytes(Convert.ToInt16(*buff)), 0, lpDstBuf, TempPos, 2);
									else
										Array.Copy(BitConverter.GetBytes(Convert.ToUInt16(*buff)), 0, lpDstBuf, TempPos, 2);
								}
								else if (MyItemColumns[k].TypeSize == 4)
								{
									if (MyItemColumns[k].Signed)
										Array.Copy(BitConverter.GetBytes(Convert.ToInt32(*buff)), 0, lpDstBuf, TempPos, 4);
									else
										Array.Copy(BitConverter.GetBytes(Convert.ToUInt32(*buff)), 0, lpDstBuf, TempPos, 4);
								}
								// увеличиваем позицию в массиве item.bmd на размер переменной
								TempX += MyItemColumns[k].TypeSize;
							}
						}
					}
				}

				// декриптование
				for (int i = 0; i < ITEM_TOTAL_LINE_SIZE; i++)
					lpDstBuf[i] ^= XorKeys[i % 3];

				using (FileStream filestream = File.Create(saveFileDialog1.FileName, ITEM_TOTAL_LINE_SIZE + 4))
				{
					// считывание потока в массив
					filestream.Write(lpDstBuf, 0, ITEM_TOTAL_LINE_SIZE);

					// Key Item.bmd for checksum
					filestream.Write(BitConverter.GetBytes(DecryptKey(lpDstBuf, ITEM_TOTAL_LINE_SIZE, 0xE2F1u)), 0, 4);
				}
			}
		}


		uint DecryptKey(byte[] pSrcBuf, int Size, uint Key)
		{
			/*	ITEM_ENG_BMD                       = $E2F1;
				SKILL_ENG_BMD                      = $5A18;
				ITEMSETTYPE_ENG_BMD                = $E5F1;
				ITEMSETOPTION_ENG_BMD              = $A2F1;
				FILTER_BMD                         = $3E7D;
				MASTERSKILLTREEDATA_BMD            = $2BC1;
				MASTERSKILLTREETOOLTIPDATA_ENG_BMD = $2BC1;
				FILTERNAME_BMD                     = $2BC1;
				ITEMTOOLTIP_BMD                    = $E2F1;
				ITEMLEVELTOOLTIP_BMD               = $E2F1;
				ITEMLEVELTOOLTIPTEXT_BMD           = $E2F1;

				FILTER_BMD_BLOCK                   = $14;*/

			uint DecryptedKey = Key << 9;
			int e = 1;
			uint result = (((DecryptedKey + BitConverter.ToUInt32(pSrcBuf, 0)) + Key) >> e) ^ (DecryptedKey + BitConverter.ToUInt32(pSrcBuf, 0));

			for (int i = 0; i < Size; )
			{
				if (i > 0)
					result = ((result + Key) >> e) ^ result;

				i += 4;
				result = result ^ BitConverter.ToUInt32(pSrcBuf, i);
				i += 4;
				result = result + BitConverter.ToUInt32(pSrcBuf, i);
				i += 4;
				result = result ^ BitConverter.ToUInt32(pSrcBuf, i);
				i += 4;
				result = result + BitConverter.ToUInt32(pSrcBuf, i);

				if (e == 1)
					e = 5;
				else
					e = 1;
			}
			return result;
		}

		// загрузка exl файла
		unsafe private bool LoadExcel(string FileName)
		{
			// открываем документ, надеемся что он структурирован этим же конвертором xD
			excelapp = new Excel.Application();
			//excelapp.Visible = true;
			Excel.Workbook excelappworkbook = excelapp.Workbooks.Open(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
																		Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

			// получаем страницы книги
			excelsheets = excelappworkbook.Worksheets;

			// сворачиваем для увеличения скорости
			excelsheets.Application.WindowState = Excel.XlWindowState.xlMinimized;
			excelsheets.Application.Visible = false;

			for (int i = 0; i < ITEM_MAX_TYPES; i++)
			{
				// поехали по страницам книги
				excelworksheet = (Excel.Worksheet)excelsheets.get_Item(i + 1);
				excelworksheet.Activate();

				progressBar3.Value = i;
				for (int j = 0; j < ITEM_MAX_IN_TYPE; j++)
				{
					// поехали по строкам
					progressBar2.Value = j;

					// считываем имя
					excelcells = (Excel.Range)excelworksheet.Cells[j + 3, 3];
					Items[i, j].ItemName = Convert.ToString(excelcells.Value2);

					// считываем остальные ячейки
					fixed (Int64* buff2 = Items[i, j].Numbers)
					{
						excelcells = excelworksheet.get_Range(ColumnTempName[3] + Convert.ToString(j + 3), ColumnTempName[3 + MyItemColumns.Length - 1] + Convert.ToString(j + 3));
						System.Array SomeNumbers = excelcells.Value2 as System.Array;

						int k = 0;
						foreach (Object Number in SomeNumbers)
						{
							Int64* buff = buff2 + k;
							*buff = Convert.ToInt64(Number);
							k++;
						}
					}
				}
			}
			excelapp.Quit();
			return true;
		}

		// структура ячейки
		public struct ExcColumn
		{
			String name;	// имя, записываемое заголовком
			String colSize;	// размер хранения, записываемый заголовком
			int width;		// ширина ячейки (в excel)
			int typeSize;	// размер переменной в байтах
			bool signed;	// переменная со знаком или без
			public ExcColumn(String N, String C, int W, int T, bool S)
			{
				name = N;
				colSize = C;
				width = W;
				typeSize = T;
				signed = S;
			}
			public string Name { get { return name; } }
			public string ColSize { get { return colSize; } }
			public int Width { get { return width; } }
			public int TypeSize { get { return typeSize; } }
			public bool Signed { get { return signed; } }
		}

		// структура ячеек Item.bmd в Excel
		ExcColumn[] MyItemColumns = new ExcColumn[]{
			new ExcColumn("Two hnd",	"0/65535",	8, 2, false),
			new ExcColumn("Item lvl",	"0/65535",	8, 2, false),
			new ExcColumn("Item slot",	"-128/127", 8, 1, true),
			new ExcColumn("Temp",		"?/?",		8, 1, true),
			new ExcColumn("Skill num",	"0/65535",	8, 2, false),
			new ExcColumn("Width",		"0/255",	6, 1, false),
			new ExcColumn("Height",		"0/255",	6, 1, false),
			new ExcColumn("< dmg",		"0/255",	6, 1, false),
			new ExcColumn("> dmg",		"0/255",	6, 1, false),
			new ExcColumn("Def rate",	"0/255",	7, 1, false),
			new ExcColumn("Defence",	"0/255",	7, 1, false),
			new ExcColumn("Magic def",	"0/255",	8, 1, false),
			new ExcColumn("Speed",		"0/255",	6, 1, false),
			new ExcColumn("Walk spd",	"0/255",	8, 1, false),
			new ExcColumn("Durab",		"0/255",	6, 1, false),
			new ExcColumn("Mag dur",	"0/255",	7, 1, false),
			new ExcColumn("Mag pow",	"0/255",	8, 1, false),
			new ExcColumn("Strength",	"0/65535",	8, 2, false),
			new ExcColumn("Agility",	"0/65535",	8, 2, false),
			new ExcColumn("Energy",		"0/65535",	8, 2, false),
			new ExcColumn("Vitality",	"0/65535",	8, 2, false),
			new ExcColumn("Command",	"0/65535",	8, 2, false),
			new ExcColumn("Req lvl",	"0/65535",	8, 2, false),
			new ExcColumn("Value",		"0/65535",	8, 2, false),
			new ExcColumn("Zen", "0/4 294 967 295",	14, 4, false),
			new ExcColumn("Type",		"0/255",	5, 1, false),
			new ExcColumn("DW",			"0/3",		5, 1, false),
			new ExcColumn("DK",			"0/3",		5, 1, false),
			new ExcColumn("Elf",		"0/3",		5, 1, false),
			new ExcColumn("MG",			"0/3",		5, 1, false),
			new ExcColumn("DL",			"0/3",		5, 1, false),
			new ExcColumn("SU",			"0/3",		5, 1, false),
			new ExcColumn("RF",			"0/3",		5, 1, false),
			new ExcColumn("Ice res",	"0/255",	8, 1, false),
			new ExcColumn("Poise res",	"0/255",	8, 1, false),
			new ExcColumn("Light res",	"0/255",	8, 1, false),
			new ExcColumn("Fire res",	"0/255",	8, 1, false),
			new ExcColumn("Earth res",	"0/255",	8, 1, false),
			new ExcColumn("Wind res",	"0/255",	8, 1, false),
			new ExcColumn("Water res",	"0/255",	8, 1, false),
			new ExcColumn("Unk",		"0/255",	8, 1, false)
		};

		// название колонок в Excel
		public String[] ColumnTempName = new String[] {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z",
														"AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ",
														"BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ"};

		// заполнение массива Item.bmd d Excel 
		unsafe private void CreateExcelItem()
		{
			// Создаем документ с 16 страницами
			excelapp = new Excel.Application();
			//excelapp.Visible=true;

			excelapp.SheetsInNewWorkbook = 1;
			Excel.Workbook excelappworkbook = excelapp.Workbooks.Add(Type.Missing);

			String[] SheetsName = new String[] { "Sword", "Axe", "MaceScepter", "Spear", "BowCrossbow", "Staff", "Shield", "Helm", "Armor", "Pants", "Gloves", "Boots", "Accessories", "Misc1", "Misc2", "Scrolls" };

			// получаем страницы книги
			excelsheets = excelappworkbook.Worksheets;

			// определяем имена страницам и переходим на страницу
			excelworksheet = (Excel.Worksheet)excelsheets.get_Item(0 + 1);
			excelworksheet.Name = SheetsName[0];
			excelworksheet.Activate();
			excelworksheet.Application.ActiveWindow.SplitColumn = 3;
			excelworksheet.Application.ActiveWindow.SplitRow = 2;
			excelworksheet.Application.ActiveWindow.FreezePanes = true;

			// заполнение Index (0.1.2.3...)
			excelcells = excelworksheet.get_Range("B3", Type.Missing);
			excelcells.Value2 = 0;
			excelcells = excelworksheet.get_Range("B4", Type.Missing);
			excelcells.Value2 = 1;
			excelcells = excelworksheet.get_Range("B3", "B4");
			Excel.Range dest = excelworksheet.get_Range("B3", "B514");
			excelcells.AutoFill(dest, Excel.XlAutoFillType.xlFillDefault);

			// сворачиваем для увеличения скорости
			excelworksheet.Application.WindowState = Excel.XlWindowState.xlMinimized;
			excelworksheet.Application.Visible = false;

			// оцентровываем первую строку
			excelcells = (Excel.Range)excelworksheet.Rows["1", Type.Missing];
			excelcells.HorizontalAlignment = Excel.Constants.xlCenter;

			// зажирняем и оцентровываем вторую строку
			excelcells = (Excel.Range)excelworksheet.Rows["2", Type.Missing];
			excelcells.Font.Bold = true;
			excelcells.HorizontalAlignment = Excel.Constants.xlCenter;

			// устанавливаем размер колонок
			excelcells = (Excel.Range)excelworksheet.Columns["A", Type.Missing];
			excelcells.ColumnWidth = 5;
			excelcells = (Excel.Range)excelworksheet.Columns["B", Type.Missing];
			excelcells.ColumnWidth = 5;
			excelcells = (Excel.Range)excelworksheet.Columns["C", Type.Missing];
			excelcells.ColumnWidth = 30;
			for (int j = 0; j < MyItemColumns.Length; j++)
			{
				excelcells = (Excel.Range)excelworksheet.Columns[ColumnTempName[j + 3], Type.Missing];
				excelcells.ColumnWidth = MyItemColumns[j].Width;
			}

			// заполняем первую строку границами как называется не помню
			excelcells = excelworksheet.get_Range("C1", Type.Missing);
			excelcells.Value2 = "Char[30]";
			excelcells.Activate();
			for (int j = 0; j < MyItemColumns.Length; j++)
			{
				excelcells = excelapp.ActiveCell.get_Offset(0, 1);
				excelcells.Value2 = MyItemColumns[j].ColSize;
				excelcells.Activate();
			}

			// заполняем вторую строку названиями
			excelcells = excelworksheet.get_Range("A2", Type.Missing);
			excelcells.Value2 = "Type";
			excelcells = excelworksheet.get_Range("B2", Type.Missing);
			excelcells.Value2 = "Index";
			excelcells = excelworksheet.get_Range("C2", Type.Missing);
			excelcells.Value2 = "Item Name";
			excelcells.Activate();
			for (int j = 0; j < MyItemColumns.Length; j++)
			{
				excelcells = excelapp.ActiveCell.get_Offset(0, 1);
				excelcells.Value2 = MyItemColumns[j].Name;
				excelcells.Activate();
			}

			// обнуляем все ячейки кроме названия
			excelcells = excelworksheet.get_Range("D3", "AR514");
			excelcells.Value2 = 0;

			// number format 12 232 232 для zen
			excelcells = excelworksheet.get_Range("AB3", "AB514");
			excelcells.NumberFormat = "# ##0";

			// копируем листы
			for (int i = 0; i < ITEM_MAX_TYPES - 1; i++)
			{
				excelworksheet = (Excel.Worksheet)excelsheets.get_Item(i + 1);
				excelworksheet.Copy(Type.Missing, excelworksheet);
				excelworksheet = (Excel.Worksheet)excelsheets.get_Item(i + 2);
				excelworksheet.Name = SheetsName[i + 1];
			}

			// заполняем ячейки
			for (int i = 0; i < ITEM_MAX_TYPES; i++)
			{
				// выделяем нужный лист
				excelworksheet = (Excel.Worksheet)excelsheets.get_Item(i + 1);
				excelworksheet.Activate();

				// заполняем тип вещей
				excelcells = excelworksheet.get_Range("A3", "A514");
				excelcells.Value2 = i;

				progressBar3.Value = i;
				// поехали по строкам
				for (int j = 0; j < ITEM_MAX_IN_TYPE; j++)
				{
					progressBar2.Value = j;

					// заполняем имя
					if (Items[i, j].ItemName[0] != '\0')
					{
						excelcells = (Excel.Range)excelworksheet.Cells[j + 3, 3];
						excelcells.Value2 = Items[i, j].ItemName;
						excelcells.Select();
					}

					// заполняем остальные ячейки
					fixed (Int64* buff = Items[i, j].Numbers)
					{
						for (int k = 0; k < MyItemColumns.Length; k++)
						{
							if (buff != null && *(buff + k) != 0.0f)
							{
								excelcells = (Excel.Range)excelworksheet.Cells[j + 3, k + 4];
								excelcells.Value2 = *(buff + k);
							}
						}
					}
				}
			}

			// показываем готовый файл
			excelapp.Visible = true;

			// перепрыгиваем на 1 лист
			excelworksheet = (Excel.Worksheet)excelsheets.get_Item(0 + 1);
			excelworksheet.Activate();

			progressBar2.Value = 0;
			progressBar3.Value = 0;
			MessageBox.Show("All Done!");
		}

		private void button4_Click(object sender, EventArgs e)
		{
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				if (openFileDialog1.FileName.EndsWith(".bmd"))
				{
					BmdFile = openFileDialog1.OpenFile();

					if (LoadItemBmd(BmdFile))
						CreateExcelItem();
					else
						MessageBox.Show("File loading error");

					BmdFile.Close();
				}
				else if (openFileDialog1.FileName.EndsWith(".xls"))
				{
					string FileName = openFileDialog1.FileName;
					if (LoadExcel(FileName))
						SaveItemBmd();
					else
						MessageBox.Show("File loading error");
					//xls load
				}
				else if (openFileDialog1.FileName.EndsWith(".xlsx"))
				{
					string FileName = openFileDialog1.FileName;
					if (LoadExcel(FileName))
						SaveItemBmd();
					else
						MessageBox.Show("File loading error");
					//xlsx load
				}
			}
		}

		private void button4_DragEnter(object sender, DragEventArgs e)
		{
			button4.Text = "Drop file here!";
			button4.FlatStyle = FlatStyle.Flat;
			if (e.Data.GetDataPresent(DataFormats.FileDrop))
			{
				e.Effect = DragDropEffects.Move | DragDropEffects.Copy | DragDropEffects.Scroll;
			}
			else
			{
				e.Effect = DragDropEffects.None;
			}
		}

		private void button4_DragLeave(object sender, EventArgs e)
		{
			button4.Text = "Load file";
			button4.FlatStyle = FlatStyle.Standard;
		}


		private void button4_DragDrop(object sender, DragEventArgs e)
		{
			button4.Text = "Load file";
			button4.FlatStyle = FlatStyle.Standard;

			string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

			foreach (string file in files)
			{
				if (!File.Exists(file))
					continue;


				if (file.EndsWith(".bmd"))
				{
					FileStream stream1 = File.Open(file, FileMode.Open);
					if (LoadItemBmd(stream1))
						CreateExcelItem();
					else
						MessageBox.Show("File loading error");

					stream1.Close();
				}
				else if (file.EndsWith(".xls"))
				{
					//xls load
				}
				else if (file.EndsWith(".xlsx"))
				{
					//xlsx load
				}
			}
		}
	}
}
