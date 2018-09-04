using System;
using System.ComponentModel;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Input;
using Xceed.Words.NET;
using Image = Xceed.Words.NET.Image;

namespace ScreenCaptureApp
{
    public class CaptureScreenViewModel : INotifyPropertyChanged
    {
        private bool _SaveAsImage;
        public bool SaveAsImage
        {
            get { return _SaveAsImage; }
            set
            {
                _SaveAsImage = value;
                OnPropertyChange("chkSaveAsImage");
                _SaveInWordPanel = _SaveAsImage ? "Collapsed" : "Visible";
                OnPropertyChange("SaveInWordPanel");
            }
        }


        private bool _SaveInWord;
        public bool SaveInWord
        {
            get { return _SaveInWord; }
            set
            {
                _SaveInWord = value;
                OnPropertyChange("chkSaveInWord");
                _SaveInWordPanel = _SaveInWord ? "Visible" : "Collapsed";
                OnPropertyChange("SaveInWordPanel");
            }
        }


        private string _SaveInWordPanel;
        public string SaveInWordPanel
        {
            get { return _SaveInWordPanel; }
            set
            {
                _SaveInWordPanel = value;
                OnPropertyChange("SaveInWordPanel");
            }
        }


        private bool _NewFile;
        public bool NewFile
        {
            get { return _NewFile; }
            set { _NewFile = value; OnPropertyChange("NewFile"); }
        }


        private bool _ExistingFile;
        public bool ExistingFile
        {
            get { return _ExistingFile; }
            set { _ExistingFile = value; OnPropertyChange("ExistingFile"); }
        }


        private string _FileName;
        public string FileName
        {
            get { return _FileName; }
            set
            {
                _FileName = value;
                OnPropertyChange("FileName");
            }
        }

        private string _ExistingFileName;

        public string ExistingFileName
        {
            get { return _ExistingFileName; }
            set { _ExistingFileName = value; OnPropertyChange("ExistingFileName"); }
        }


        public ICommand BrowseButton_Click { get; private set; }
        public ICommand WindowClosing { get; private set; }

        ScreenCapture sc;
        Thread keyloggerThread;
        private string WordFileDirectory;
        private string ImageFileDirectory;
        private bool flag = true;

        public CaptureScreenViewModel()
        {
            WordFileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\Screen Capture\\";
            ImageFileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "\\Screen Capture\\Image\\" + DateTime.Now.ToString("dd-MM-yyyy") + "\\";
            _FileName = DateTime.Now.ToString("dd-MM-yyyy");
            _SaveInWordPanel = "Collapsed";
            _SaveAsImage = true;
            _NewFile = true;

            ThreadStart();

            BrowseButton_Click = new RelayCommand(BrowseButtonClick);
            WindowClosing = new RelayCommand(WindowClosingEvent);
        }

        public void StartMonitor()
        {
            DataSource dataSource = new DataSource();
            while (true)
            {
                // Get pressed keys and saves them
                if (dataSource.IsKeyPress(Key.PrintScreen))
                {
                    if (_SaveAsImage)
                    {
                        SaveToImage();
                    }
                    else if (_SaveInWord)
                    {
                        if (_NewFile)
                        {
                            SaveToNewFile();
                        }
                        else if (_ExistingFile)
                        {
                            SaveToExistingFile();
                        }
                    }
                }

            }
        }

        private void SaveToImage()
        {
            try
            {

                if (!Directory.Exists(ImageFileDirectory))
                {
                    Directory.CreateDirectory(ImageFileDirectory);
                }
                ScreenCapture sc = new ScreenCapture();
                // capture entire screen, and save it to a file
                System.Drawing.Image img = sc.CaptureScreen();

                // capture this window, and save it;
                //sc.CaptureWindowToFile(this.Handle, Guid.NewGuid()+".gif", ImageFormat.Gif);
                img.Save(ImageFileDirectory + "\\" + "Captured-" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".jpg", ImageFormat.Jpeg);

            }
            catch (Exception ex)
            {
                flag = true;
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ThreadStart();
            }
        }

        private void SaveToNewFile()
        {
            try
            {
                string fileName = !string.IsNullOrEmpty(_FileName) ? _FileName + ".docx" : DateTime.Now.ToString("dd-MM-yyyy") + ".docx";

                if (!IsFileLocked(WordFileDirectory + fileName))
                {
                    sc = new ScreenCapture();

                    // capture entire screen, and save it to a file
                    System.Drawing.Image img1 = sc.CaptureScreen();

                    // capture this window, and save it;
                    //sc.CaptureWindowToFile(this.Handle, Guid.NewGuid()+".gif", ImageFormat.Gif);
                    img1.Save("Screenshot" + ".jpg", ImageFormat.Jpeg);

                    if (!Directory.Exists(WordFileDirectory))
                    {
                        Directory.CreateDirectory(WordFileDirectory);
                    }

                    if (!File.Exists(WordFileDirectory + fileName))
                    {
                        var docCreate = DocX.Create(WordFileDirectory + fileName);
                        docCreate.Save();
                    }
                    var docload = DocX.Load(WordFileDirectory + fileName);
                    Image img = docload.AddImage("Screenshot.jpg");
                    Picture p = img.CreatePicture();
                    p.Width = 690;
                    p.Height = 365;

                    Paragraph par = docload.InsertParagraph(Convert.ToString(" "));
                    par.AppendPicture(p);
                    docload.Save();
                }
                else
                {
                    MessageBox.Show("The action can't be completed because the file is open in System.", "Error", MessageBoxButtons.OK);
                }
            }
            catch (IOException ex)
            {
                flag = true;
                MessageBox.Show("The action can't be completed because the file is open in System.", "Error", MessageBoxButtons.OK);
            }
            catch (Exception ex)
            {
                flag = true;
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
            }
            finally
            {
                ThreadStart();
            }
        }

        private void SaveToExistingFile()
        {
            try
            {
                sc = new ScreenCapture();

                // capture entire screen, and save it to a file
                System.Drawing.Image img1 = sc.CaptureScreen();

                // capture this window, and save it;
                //sc.CaptureWindowToFile(this.Handle, Guid.NewGuid()+".gif", ImageFormat.Gif);
                img1.Save("Screenshot" + ".jpg", ImageFormat.Jpeg);

                if (!File.Exists(ExistingFileName))
                {
                    var docCreate = DocX.Create(ExistingFileName);
                    docCreate.Save();
                }
                var docload = DocX.Load(ExistingFileName);
                Image img = docload.AddImage("Screenshot.jpg");
                Picture p = img.CreatePicture();
                p.Width = 690;
                p.Height = 365;

                Paragraph par = docload.InsertParagraph(Convert.ToString(" "));
                par.AppendPicture(p);
                docload.Save();
            }
            catch (IOException ex)
            {
                flag = true;
                MessageBox.Show("The action can't be completed because the file is open in System.", "Error", MessageBoxButtons.OK);
            }
            catch (Exception ex)
            {
                flag = true;
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
            }
            finally
            {
                ThreadStart();
            }
        }

        private void BrowseButtonClick(object obj)
        {
            try
            {
                // Create OpenFileDialog
                Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();

                // Launch OpenFileDialog by calling ShowDialog method
                Nullable<bool> result = openFileDlg.ShowDialog();

                // Get the selected file name and display in a TextBox.
                // Load content of file in a TextBlock
                if (result == true)
                {
                    if (openFileDlg.SafeFileNames.Count() == 1)
                    {
                        if (openFileDlg.SafeFileName.Contains(".doc"))
                        {
                            ExistingFileName = openFileDlg.FileName;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                flag = true;
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
            }
            finally
            {
                ThreadStart();
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChange(string propertyname)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyname));
            }
        }

        private void WindowClosingEvent(object obj)
        {
            keyloggerThread.Abort();
        }

        private void ThreadStart()
        {

            if (flag)
            {
                // Start keylogger in another thread and close GUI
                keyloggerThread = new Thread(() =>
                {
                    StartMonitor();
                });
                keyloggerThread.SetApartmentState(ApartmentState.STA);
                keyloggerThread.Name = "KeyloggerThread";
                keyloggerThread.Start();

                flag = false;
            }
        }

        private bool IsFileLocked(string file)
        {
            try
            {
                using (var stream = File.OpenRead(file))
                    return false;
            }
            catch (IOException)
            {
                flag = true;
                return true;
            }
        }
    }
}
