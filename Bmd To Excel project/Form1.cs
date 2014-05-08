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

        private Excel.Application excelapp;
		private Excel.Window excelWindow; 
		
		private Excel.Sheets excelsheets;
		private Excel.Worksheet excelworksheet;

		private Excel.Range excelcells;
				
		Stream BmdFile; 

		const int ITEM_LINE_SIZE = 84;

		unsafe public struct ItemEx
		{
			public string ItemName;
			public fixed Int64 Numbers[54];
		};

		public ItemEx[,] Items = new ItemEx[16,512];

		unsafe private bool LoadItemBmd(Stream File)
        {
			const int ITEM_TOTAL_LINE_SIZE = 512 * 16 * 84; //688128
			
			byte[] lpDstBuf = new byte[ITEM_TOTAL_LINE_SIZE];						
			byte[] XorKeys = new byte[3] { 0xFC, 0xCF, 0xAB };

			for (int i = 0; i < ITEM_TOTAL_LINE_SIZE; i++)
			{
				lpDstBuf[i] = 0;
			}
			File.Read(lpDstBuf, 0, ITEM_TOTAL_LINE_SIZE);

			// decription
			for (int i = 0; i < ITEM_TOTAL_LINE_SIZE; i++)
			{
				lpDstBuf[i] ^= XorKeys[i % 3];
			}

			for (int i = 0; i < 16; i++)
			{
				for (int j = 0; j < 512; j++)
				{
					int start = (i * 512 + j)*84;
					Items[i, j].ItemName = System.Text.Encoding.Default.GetString(lpDstBuf, start, 30);
					fixed (Int64* buff2 = Items[i, j].Numbers)
					{
						int TempX = 0;
						for (int k = 0; k < MyItemColumns.Length; k++)
						{
							Int64* buff = buff2 + k;
							
							if (MyItemColumns[k].TypeSize == 1)
							{
								if (MyItemColumns[k].Signed == 1)
									*buff = (sbyte)lpDstBuf[start + 30 + TempX];
								else
									*buff = (byte)lpDstBuf[start + 30 + TempX];
							}
							else if (MyItemColumns[k].TypeSize == 2)
							{
								if (MyItemColumns[k].Signed == 1)
									*buff = BitConverter.ToInt16(lpDstBuf, start + 30 + TempX);
								else
									*buff = BitConverter.ToUInt16(lpDstBuf, start + 30 + TempX);
							}
							else if (MyItemColumns[k].TypeSize == 4)
							{
								if (MyItemColumns[k].Signed == 1)
									*buff = BitConverter.ToInt32(lpDstBuf, start + 30 + TempX);
								else
									*buff = BitConverter.ToUInt32(lpDstBuf, start + 30 + TempX);
							}
							TempX += MyItemColumns[k].TypeSize;
						}
					}
				}
			}
			return true;
        }

        private bool LoadExcel()
        {
			return true;
        }

		public struct ExcColumn
		{
			String name;
			String colSize;
			int width;
			int typeSize;
			int signed;
			public ExcColumn(String N, String C, int W, int T, int S)
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
			public int Signed { get { return signed; } }

		}
		ExcColumn[] MyItemColumns = new ExcColumn[]{
			new ExcColumn("Two hnd",	"0/65535",	8, 2, 0),
			new ExcColumn("Item lvl",	"0/65535",	8, 2, 0),
			new ExcColumn("Item slot",	"-128/127", 8, 1, 1),
			new ExcColumn("Temp",		"?/?",		8, 1, 1),
			new ExcColumn("Skill num",	"0/65535",	8, 2, 0),
			new ExcColumn("Width",		"0/255",	6, 1, 0),
			new ExcColumn("Height",		"0/255",	6, 1, 0),
			new ExcColumn("< dmg",		"0/255",	6, 1, 0),
			new ExcColumn("> dmg",		"0/255",	6, 1, 0),
			new ExcColumn("Def rate",	"0/255",	7, 1, 0),
			new ExcColumn("Defence",	"0/255",	7, 1, 0),
			new ExcColumn("Magic def",	"0/255",	8, 1, 0),
			new ExcColumn("Speed",		"0/255",	6, 1, 0),
			new ExcColumn("Walk spd",	"0/255",	8, 1, 0),
			new ExcColumn("Durab",		"0/255",	6, 1, 0),
			new ExcColumn("Mag dur",	"0/255",	7, 1, 0),
			new ExcColumn("Mag pow",	"0/255",	8, 1, 0),
			new ExcColumn("Strength",	"0/65535",	8, 2, 0),
			new ExcColumn("Agility",	"0/65535",	8, 2, 0),
			new ExcColumn("Energy",		"0/65535",	8, 2, 0),
			new ExcColumn("Vitality",	"0/65535",	8, 2, 0),
			new ExcColumn("Command",	"0/65535",	8, 2, 0),
			new ExcColumn("Req lvl",	"0/65535",	8, 2, 0),
			new ExcColumn("Value",		"0/65535",	8, 2, 0),
			new ExcColumn("Zen", "0/4 294 967 295",	14, 4, 0),
			new ExcColumn("Type",		"0/255",	5, 1, 0),
			new ExcColumn("DW",			"0/3",		5, 1, 0),
			new ExcColumn("DK",			"0/3",		5, 1, 0),
			new ExcColumn("Elf",		"0/3",		5, 1, 0),
			new ExcColumn("MG",			"0/3",		5, 1, 0),
			new ExcColumn("DL",			"0/3",		5, 1, 0),
			new ExcColumn("SU",			"0/3",		5, 1, 0),
			new ExcColumn("RF",			"0/3",		5, 1, 0),
			new ExcColumn("Ice res",	"0/255",	8, 1, 0),
			new ExcColumn("Poise res",	"0/255",	8, 1, 0),
			new ExcColumn("Light res",	"0/255",	8, 1, 0),
			new ExcColumn("Fire res",	"0/255",	8, 1, 0),
			new ExcColumn("Earth res",	"0/255",	8, 1, 0),
			new ExcColumn("Wind res",	"0/255",	8, 1, 0),
			new ExcColumn("Water res",	"0/255",	8, 1, 0),
			new ExcColumn("Unk",		"0/255",	8, 1, 0)
		};

		public String[] ColumnTempName = new String[] {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z",
														"AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ",
														"BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ"};

		unsafe private void CreateExcelItem()
		{
			// Создаем документ с 16 страницами
            excelapp = new Excel.Application(); 
            //excelapp.Visible=true;

            excelapp.SheetsInNewWorkbook=1;
            Excel.Workbook excelappworkbook = excelapp.Workbooks.Add(Type.Missing);

			String[] SheetsName = new String[16] { "Sword", "Axe", "MaceScepter", "Spear", "BowCrossbow", "Staff", "Shield", "Helm", "Armor", "Pants", "Gloves", "Boots", "Accessories", "Misc1", "Misc2", "Scrolls" };

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
			for (int i = 0; i < 15; i++)
			{
				excelworksheet = (Excel.Worksheet)excelsheets.get_Item(i + 1);
				excelworksheet.Copy(Type.Missing, excelworksheet);
				excelworksheet = (Excel.Worksheet)excelsheets.get_Item(i + 2);
				excelworksheet.Name = SheetsName[i + 1];
			}

			// заполняем ячейки
			for (int i = 0; i < 16; i++)
			{
				// выделяем нужный лист
				excelworksheet = (Excel.Worksheet)excelsheets.get_Item(i + 1);
				excelworksheet.Activate();

				excelcells = excelworksheet.get_Range("A3", "A514");
				excelcells.Value2 = i;

				progressBar3.Value = i;
				// поехали по строкам
				for (int j = 0; j < 512; j++)
				{
					progressBar2.Value = j;
					if (Items[i,j].ItemName[0] != '\0')
					{
						excelcells = (Excel.Range)excelworksheet.Cells[j + 3, 3];
						excelcells.Value2 = Items[i, j].ItemName;
						excelcells.Select();
					}
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
					//xls load
				}
				else if (openFileDialog1.FileName.EndsWith(".xlsx"))
				{
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
