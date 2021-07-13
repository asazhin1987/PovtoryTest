using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Comparer.DTO;
using ExcelComparer.ISvc;
using Microsoft.Win32;
using Povtory.Commands;
using Povtory.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Povtory.ViewModels
{
    class BLViewModel : INotifyPropertyChanged
    {
		private readonly IComparerSvc svc;
		public BLViewModel(IComparerSvc _svc)
        {
			svc = _svc;
		}


		void aaa()
		{
			ResultDTO result = svc.Compaere();

			Otstuplenie otstuplenie = new OtstuplenieNeud()
			{
				DateOfInspection = DateTime.Now,
				DistanciaPuti = 1,
				DlinaNeispravnosti = 1
			};
		}
	//command CompareCommand

/*====*/


        private List<List<Otstuplenie>> newFile;
        private List<List<Otstuplenie>> oldFile;
        private List<Otstuplenie> povtorNeudList;
        private List<Otstuplenie> povtorStep3List;
        private List<Otstuplenie> povtorStep2k3List;

        public List<List<Otstuplenie>> NewFile
        {
            get
            {
                return newFile;
            }
            set
            {
                newFile = value;
            }
        }

        public List<List<Otstuplenie>> OldFile
        {
            get
            {
                return oldFile;
            }
            set
            {
                oldFile = value;
            }
        }
        public List<Otstuplenie> PovtorNeudList
        {
            get
            {
                return povtorNeudList;
            }
            set
            {
                povtorNeudList = value;
            }
        }
        public List<Otstuplenie> PovtorStep3List
        {
            get
            {
                return povtorStep3List;
            }
            set
            {
                povtorStep3List = value;
            }
        }
        public List<Otstuplenie> PovtorStep2k3List
        {
            get
            {
                return povtorStep2k3List;
            }
            set
            {
                povtorStep2k3List = value;
            }
        }

        //Путь к файлу последнего проезда
        private string newFilePath;
        //Путь к файлу прошлого проезда
        private string oldFilePath;

        public string NewFilePath
        {
            get
            {
                return newFilePath;
            }
            set
            {
                newFilePath = value;
                OnPropertyChanged();
            }
        }

        public string OldFilePath
        {
            get
            {
                return oldFilePath;
            }
            set
            {
                oldFilePath = value;
                OnPropertyChanged();
            }
        }

        #region Реализация интерфейса INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        internal void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        //Метод открытия окна выбора файла
        public void OpenFileDialogMethod(object sender)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Выберите файл";
            openFileDialog.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            openFileDialog.DefaultExt = "*.xls;*.xlsx";
            openFileDialog.Multiselect = false;
            //openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments

            if (openFileDialog.ShowDialog() == true)
            {
                if (((Button)sender).Name == "btnNewOFD")
                {
                    NewFilePath = openFileDialog.FileName;
                }
                else if (((Button)sender).Name == "btnOldOFD")
                {
                    OldFilePath = openFileDialog.FileName;
                }
            }
        }

        //Количество строк в файле
        private int countRecordsInFile;
        public int CountRecordsInFile
        {
            get
            {
                return countRecordsInFile;
            }
            set
            {
                countRecordsInFile = value;
                OnPropertyChanged();
            }
        }

        //Метод для выгрузки данных из всех трех листов файла проезда в List
        private List<List<Otstuplenie>> CreateFullFileListOfNewOtstupleniy(string filePath)
        {
            List<List<Otstuplenie>> fullListNew = new List<List<Otstuplenie>>();

            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    throw new ArgumentNullException();
                }

                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException();
                }

                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbookNewFile = excel.Workbooks.Open(filePath);

                //Перебор по всем трем вкладкам в файле нового проезда
                for (int i = 0; i < 3; i++)
                {
                    Excel.Worksheet sheetNewFile = workbookNewFile.Worksheets[i + 1];

                    // Ищем последнюю заполненную ячейку в файле нового проезда
                    Excel.Range FileRangeNewFile = sheetNewFile.UsedRange;
                    CountRecordsInFile = FileRangeNewFile.Rows.Count;


                    if (i == 0)
                    {
                        List<Otstuplenie> listNeud = new List<Otstuplenie>();
                        for (int j = 5; j <= CountRecordsInFile; j++)
                        {
                            //if (sheetNewFile.Cells[j, 10].Value == 3 || sheetNewFile.Cells[j, 10].Value == 4)
                            //{
                            OtstuplenieNeud neud = new OtstuplenieNeud(Convert.ToInt32(sheetNewFile.Cells[j, 1].Value),
                                                                       Convert.ToDateTime(sheetNewFile.Cells[j, 2].Value),
                                                                       Convert.ToByte(sheetNewFile.Cells[j, 3].Value),
                                                                       Convert.ToByte(sheetNewFile.Cells[j, 4].Value),
                                                                       Convert.ToString(sheetNewFile.Cells[j, 5].Value),
                                                                       Convert.ToString(sheetNewFile.Cells[j, 6].Value),
                                                                       Convert.ToInt32(sheetNewFile.Cells[j, 7].Value),
                                                                       Convert.ToInt32(sheetNewFile.Cells[j, 8].Value),
                                                                       Convert.ToInt32(sheetNewFile.Cells[j, 9].Value),
                                                                       Convert.ToString(sheetNewFile.Cells[j, 10].Value),
                                                                       Convert.ToString(sheetNewFile.Cells[j, 11].Value),
                                                                       Convert.ToDouble(sheetNewFile.Cells[j, 12].Value),
                                                                       Convert.ToInt32(sheetNewFile.Cells[j, 13].Value),
                                                                       Convert.ToString(sheetNewFile.Cells[j, 15].Value),
                                                                       Convert.ToString(sheetNewFile.Cells[j, 17].Value),
                                                                       Convert.ToString(sheetNewFile.Cells[j, 14].Value),
                                                                       Convert.ToString(sheetNewFile.Cells[j, 16].Value));

							OtstuplenieNeud AAAAA = new OtstuplenieNeud()
							{
								 DlinaNeispravnosti = Convert.ToString(sheetNewFile.Cells[j, 11].Value), 
								 DistanciaPuti = Convert.ToString(sheetNewFile.Cells[j, 10].Value),
								 SpeedReduction = Convert.ToString(sheetNewFile.Cells[j, 17].Value),
							};

							listNeud.Add(neud);
                            //}
                        }
                        fullListNew.Add(listNeud);
                    }
                    else
                    {
                        List<Otstuplenie> listObychnoe = new List<Otstuplenie>();

                        for (int j = 4; j <= CountRecordsInFile; j++)
                        {
                            OtstuplenueObychnoe obychn = new OtstuplenueObychnoe(Convert.ToInt32(sheetNewFile.Cells[j, 1].Value),
                                                                                 Convert.ToDateTime(sheetNewFile.Cells[j, 2].Value),
                                                                                 Convert.ToByte(sheetNewFile.Cells[j, 3].Value),
                                                                                 Convert.ToByte(sheetNewFile.Cells[j, 4].Value),
                                                                                 Convert.ToString(sheetNewFile.Cells[j, 5].Value),
                                                                                 Convert.ToString(sheetNewFile.Cells[j, 6].Value),
                                                                                 Convert.ToInt32(sheetNewFile.Cells[j, 7].Value),
                                                                                 Convert.ToInt32(sheetNewFile.Cells[j, 8].Value),
                                                                                 Convert.ToInt32(sheetNewFile.Cells[j, 9].Value),
                                                                                 Convert.ToString(sheetNewFile.Cells[j, 10].Value),
                                                                                 Convert.ToString(sheetNewFile.Cells[j, 11].Value),
                                                                                 Convert.ToDouble(sheetNewFile.Cells[j, 12].Value),
                                                                                 Convert.ToInt32(sheetNewFile.Cells[j, 13].Value),
                                                                                 Convert.ToInt32(sheetNewFile.Cells[j, 14].Value),
                                                                                 Convert.ToString(sheetNewFile.Cells[j, 15].Value),
                                                                                 Convert.ToString(sheetNewFile.Cells[j, 16].Value));

                            listObychnoe.Add(obychn);
                        }
                        fullListNew.Add(listObychnoe);
                    }
                }
                workbookNewFile.Close(true, Type.Missing, Type.Missing);
            }

            catch(ArgumentNullException nullEx)
            {
                MessageBox.Show("Путь выбора файла не заполнен", "Ошибка", MessageBoxButton.OK);
                return null;
            }
            catch(FileNotFoundException notFoundEx)
            {
                MessageBox.Show($"Указанный файл не существует:\n {filePath}", "Ошибка", MessageBoxButton.OK);
                return null;
            }
            catch(Exception e)
            {
                MessageBox.Show($"Не корректный формат данных в файле:\n {filePath}","Ошибка", MessageBoxButton.OK);
                return null;
            }
            
            return fullListNew;
        }
        
        // Метод сравнения Неудов
        public List<Otstuplenie> CompareMethodNeudVsNeud()
        {
            List<Otstuplenie> povtoryListNeud = new List<Otstuplenie>();

            foreach(OtstuplenieNeud itemInVkladkaNew in newFile[0])
            {
                foreach(OtstuplenieNeud itemInVkladkaOld in oldFile[0])
                {
                    if (itemInVkladkaOld.Neispravnost == itemInVkladkaNew.Neispravnost &&
                       itemInVkladkaOld.KmCoord == itemInVkladkaNew.KmCoord &&
                       ((itemInVkladkaNew.MetrCoord >= (itemInVkladkaOld.MetrCoord - 20)) && (itemInVkladkaNew.MetrCoord <= (itemInVkladkaOld.MetrCoord + 20))))
                    {
                        povtoryListNeud.Add(itemInVkladkaNew);
                        povtoryListNeud.Add(itemInVkladkaOld);
                    }
                }
            }
            return povtoryListNeud;
        }

        //Метод сравнения Степеней Третьих
        public List<Otstuplenie> CompareMethodStepen3VsStepen3()
        {
            List<Otstuplenie> povtoryListStep3 = new List<Otstuplenie>();

            foreach (OtstuplenueObychnoe itemInVkladkaNew in newFile[1])
            {
                foreach (OtstuplenueObychnoe itemInVkladkaOld in oldFile[1])
                {
                    if (itemInVkladkaOld.Neispravnost == itemInVkladkaNew.Neispravnost &&
                       itemInVkladkaOld.KmCoord == itemInVkladkaNew.KmCoord &&
                       ((itemInVkladkaNew.MetrCoord >= (itemInVkladkaOld.MetrCoord - 20)) && (itemInVkladkaNew.MetrCoord <= (itemInVkladkaOld.MetrCoord + 20))))
                    {
                        povtoryListStep3.Add(itemInVkladkaNew);
                        povtoryListStep3.Add(itemInVkladkaOld);
                    }
                }
            }
            return povtoryListStep3;
        }

        //Метод сравнения Степеней Вторых к Третьим
        public List<Otstuplenie> CompareMethodStepen2k3VsStepen2k3()
        {
            List<Otstuplenie> povtoryListStep3 = new List<Otstuplenie>();

            foreach (OtstuplenueObychnoe itemInVkladkaNew in newFile[2])
            {
                foreach (OtstuplenueObychnoe itemInVkladkaOld in oldFile[2])
                {
                    if (itemInVkladkaOld.Neispravnost == itemInVkladkaNew.Neispravnost &&
                       itemInVkladkaOld.KmCoord == itemInVkladkaNew.KmCoord &&
                       ((itemInVkladkaNew.MetrCoord >= (itemInVkladkaOld.MetrCoord - 20)) && (itemInVkladkaNew.MetrCoord <= (itemInVkladkaOld.MetrCoord + 20))))
                    {
                        povtoryListStep3.Add(itemInVkladkaNew);
                        povtoryListStep3.Add(itemInVkladkaOld);
                    }
                }
            }
            return povtoryListStep3;
        }

        public async void ExecuteAsync()
        {
            await Task.Run(() => Execute());
        }
        public void Execute()
        {
            NewFile = CreateFullFileListOfNewOtstupleniy(newFilePath);
            if(NewFile != null)
            {
                OldFile = CreateFullFileListOfNewOtstupleniy(oldFilePath);
                if(OldFile != null)
                {
                    PovtorNeudList = CompareMethodNeudVsNeud();
                    PovtorStep3List = CompareMethodStepen3VsStepen3();
                    PovtorStep2k3List = CompareMethodStepen2k3VsStepen2k3();
                    PrintNewExcelFilePovtoryNeud(PovtorNeudList);
                    PrintNewExcelFilePovtoryStepenAsync(PovtorStep3List, 3);
                    PrintNewExcelFilePovtoryStepenAsync(PovtorStep2k3List, 2);
                }
            }
        }

        #region Методы заполнения таблиц Excel новыми (обработанными) данными
        public async void PrintNewPrintNewExcelFilePovtoryNeudAsync(List<Otstuplenie> neudList)
        {
            await Task.Run(() => PrintNewExcelFilePovtoryNeud(neudList));
        }
        //Метод заполнения таблицы Excel неудами
        public void PrintNewExcelFilePovtoryNeud(List<Otstuplenie> neudList)
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                app = new Excel.Application();
                app.Visible = false;
                workbook = app.Workbooks.Add(1);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
            }
            catch (Exception e)
            {
                Console.Write("Error");
            }
            worksheet.Cells[1, 1].Value = "Таблица неудов";
            worksheet.Cells[2, 1].Value = "ПС";
            worksheet.Cells[2, 2].Value = "Дата проверки";
            worksheet.Cells[2, 3].Value = "ПЧ";
            worksheet.Cells[2, 4].Value = "Околоток";
            worksheet.Cells[2, 5].Value = "Участок";
            worksheet.Cells[2, 6].Value = "Путь";
            worksheet.Cells[2, 7].Value = "Км";
            worksheet.Cells[2, 8].Value = "Пк";
            worksheet.Cells[2, 9].Value = "Метр";
            worksheet.Cells[2, 10].Value = "Степень";
            worksheet.Cells[2, 11].Value = "Неисправность";
            worksheet.Cells[2, 12].Value = "Величина неисправности";
            worksheet.Cells[2, 13].Value = "Длина неисправности";
            worksheet.Cells[2, 14].Value = "Ограничение скорости км/ч";
            worksheet.Cells[2, 15].Value = "Повторы (раз)";
            worksheet.Cells[2, 16].Value = "Время ограничения";
            worksheet.Cells[2, 17].Value = "Вид проверки";

            for (int i = 0; i < neudList.Count; i++)
            {
                worksheet.Cells[i + 3, 1].Value = ((OtstuplenieNeud)(neudList[i])).PSnumber;
                worksheet.Cells[i + 3, 2].Value = ((OtstuplenieNeud)(neudList[i])).DateOfInspection;
                worksheet.Cells[i + 3, 3].Value = ((OtstuplenieNeud)(neudList[i])).DistanciaPuti;
                worksheet.Cells[i + 3, 4].Value = ((OtstuplenieNeud)(neudList[i])).Okolotok;
                worksheet.Cells[i + 3, 5].Value = ((OtstuplenieNeud)(neudList[i])).Uchastok;
                worksheet.Cells[i + 3, 6].Value = ((OtstuplenieNeud)(neudList[i])).WayNumber;
                worksheet.Cells[i + 3, 7].Value = ((OtstuplenieNeud)(neudList[i])).KmCoord;
                worksheet.Cells[i + 3, 8].Value = ((OtstuplenieNeud)(neudList[i])).PkCoord;
                worksheet.Cells[i + 3, 9].Value = ((OtstuplenieNeud)(neudList[i])).MetrCoord;
                worksheet.Cells[i + 3, 10].Value = ((OtstuplenieNeud)(neudList[i])).Stepen;
                worksheet.Cells[i + 3, 11].Value = ((OtstuplenieNeud)(neudList[i])).Neispravnost;
                worksheet.Cells[i + 3, 12].Value = ((OtstuplenieNeud)(neudList[i])).VelichinaNeispravnosti;
                worksheet.Cells[i + 3, 13].Value = ((OtstuplenieNeud)(neudList[i])).DlinaNeispravnosti;
                worksheet.Cells[i + 3, 14].Value = ((OtstuplenieNeud)(neudList[i])).SpeedReduction;
                worksheet.Cells[i + 3, 15].Value = ((OtstuplenieNeud)(neudList[i])).Povtory;
                worksheet.Cells[i + 3, 16].Value = ((OtstuplenieNeud)(neudList[i])).TimeOfRestriction;
                worksheet.Cells[i + 3, 17].Value = ((OtstuplenieNeud)(neudList[i])).VidProverki;
            }
            app.Visible = true;
        }

        public async void PrintNewExcelFilePovtoryStepenAsync(List<Otstuplenie> neudList, int stepen)
        {
            await Task.Run(() => PrintNewExcelFilePovtoryStepen(neudList, stepen));
        }
        //Метод заполения таблицы Excel степенями 3 и 2к3 (один и тот же метод)
        private void PrintNewExcelFilePovtoryStepen(List<Otstuplenie> newPovtorList, int stepen)
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                app = new Excel.Application();
                app.Visible = false;
                workbook = app.Workbooks.Add(1);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
            }
            catch (Exception e)
            {
                Console.Write("Error");
            }
            if(stepen == 3)
            {
                worksheet.Cells[1, 1].Value = "Таблица сравнения Третьих степеней";
            }
            else if(stepen == 2)
            {
                worksheet.Cells[1, 1].Value = "Таблица сравнения Вторых степеней к Третьим";
            }

            worksheet.Cells[2, 1].Value = "ПС";
            worksheet.Cells[2, 2].Value = "Дата проверки";
            worksheet.Cells[2, 3].Value = "ПЧ";
            worksheet.Cells[2, 4].Value = "Околоток";
            worksheet.Cells[2, 5].Value = "Участок";
            worksheet.Cells[2, 6].Value = "Путь";
            worksheet.Cells[2, 7].Value = "Км";
            worksheet.Cells[2, 8].Value = "Пк";
            worksheet.Cells[2, 9].Value = "Метр";
            worksheet.Cells[2, 10].Value = "Степень";
            worksheet.Cells[2, 11].Value = "Неисправность";
            worksheet.Cells[2, 12].Value = "Величина неисправности";
            worksheet.Cells[2, 13].Value = "Длина неисправности";
            worksheet.Cells[2, 14].Value = "Штук";
            worksheet.Cells[2, 15].Value = "Повторы";
            worksheet.Cells[2, 16].Value = "Вид проверки";

            for (int i = 0; i < newPovtorList.Count; i++)
            {
                worksheet.Cells[i + 3, 1].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).PSnumber;
                worksheet.Cells[i + 3, 2].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).DateOfInspection;
                worksheet.Cells[i + 3, 3].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).DistanciaPuti;
                worksheet.Cells[i + 3, 4].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).Okolotok;
                worksheet.Cells[i + 3, 5].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).Uchastok;
                worksheet.Cells[i + 3, 6].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).WayNumber;
                worksheet.Cells[i + 3, 7].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).KmCoord;
                worksheet.Cells[i + 3, 8].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).PkCoord;
                worksheet.Cells[i + 3, 9].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).MetrCoord;
                worksheet.Cells[i + 3, 10].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).Stepen;
                worksheet.Cells[i + 3, 11].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).Neispravnost;
                worksheet.Cells[i + 3, 12].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).VelichinaNeispravnosti;
                worksheet.Cells[i + 3, 13].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).DlinaNeispravnosti;
                worksheet.Cells[i + 3, 14].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).Shtuk;
                worksheet.Cells[i + 3, 15].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).Povtory;
                worksheet.Cells[i + 3, 16].Value = ((OtstuplenueObychnoe)(newPovtorList[i])).VidProverki;
            }
            app.Visible = true;
        }
        #endregion

        #region Commands

        #region Команда для запуска сравнения файлов
        // Команда Сравнения
        public ICommand OneExecuteCommand 
        {
            get
            {
                return new RelayCommand(OnExecuteCommandExecuted, CanExecuteCommandExecute);
            }
        }
        private async void OnExecuteCommandExecuted(object p)
        {
            ((ProgressBar)p).Visibility = Visibility.Visible;
            await Task.Run(() =>
            {
                
                Execute();
            });
            ((ProgressBar)p).Visibility = Visibility.Hidden;
        }
        private bool CanExecuteCommandExecute(object p)
        {
            if (p == null)
            {
                return true;
            }
            else
            {
                if (((ProgressBar)p).Visibility == Visibility.Hidden)
                {

                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
        #endregion

        #region Команда Отрытия меню выбора файла
        //Команда Открытия FileDialog
        public ICommand OpenFileDialogCommand
        {
            get
            {
                return new RelayCommand(OnOpenFileDialogCommandExecuted, CanOpeFileDialogCommandExecuted);
            }
        }
        private void OnOpenFileDialogCommandExecuted(object p)
        {
            OpenFileDialogMethod(p);
        }
        private bool CanOpeFileDialogCommandExecuted(object p)
        {
            return true;
        }
        #endregion

        #endregion

        #region ToDo
        //public List<Otstuplenie> CompareMethodNeudKilometr()
        //{
        //    List<Otstuplenie> neudKmList = new List<Otstuplenie>();

        //    foreach (OtstuplenieNeud itemInVkladkaNew in newFile[0])
        //    {
        //        foreach (OtstuplenieNeud itemInVkladkaOld in oldFile[0])
        //        {
        //            if (itemInVkladkaOld.Neispravnost == itemInVkladkaNew.Neispravnost &&
        //               itemInVkladkaOld.KmCoord == itemInVkladkaNew.KmCoord &&
        //               ((itemInVkladkaNew.MetrCoord >= (itemInVkladkaOld.MetrCoord - 20)) && (itemInVkladkaNew.MetrCoord <= (itemInVkladkaOld.MetrCoord + 20))))
        //            {
        //                neudKmList.Add(itemInVkladkaNew);
        //                neudKmList.Add(itemInVkladkaOld);
        //            }
        //        }
        //    }
        //    return neudKmList;
        //}
        #endregion

    }
}

