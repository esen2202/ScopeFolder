using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using EsenClassLib;

namespace ScopeFolder
{
    static class Program
    {
        #region genel tanımlamalar
        public static string filePath, folderPath; // sağ tıklanan dosyanın yolunu tutacak
        #endregion


        #region Klasörü Taşı Eski Version (Kopyalanan Root Klasörü Oluşturmuyor)
        public static void MoveDirectory(string source, string target)
        {
            var sourcePath = source.TrimEnd('\\', ' ');
            var targetPath = target.TrimEnd('\\', ' ');
            var lastDirectory = Path.GetFileName(sourcePath);
            targetPath = targetPath + "\\" + lastDirectory;

            var stack = new Stack<Folders>();
            stack.Push(new Folders(sourcePath, targetPath));

            while (stack.Count > 0)
            {
                var folders = stack.Pop();
                Directory.CreateDirectory(folders.Target);
                foreach (var file in Directory.GetFiles(folders.Source, "*.*"))
                {
                    string targetFile = Path.Combine(folders.Target, Path.GetFileName(file));
                    if (File.Exists(targetFile)) File.Delete(targetFile);
                    File.Move(file, targetFile);
                }

                foreach (var folder in Directory.GetDirectories(folders.Source))
                {
                    stack.Push(new Folders(folder, Path.Combine(folders.Target, Path.GetFileName(folder))));
                }
            }
            Directory.Delete(sourcePath, true);
        }

        public class Folders
        {
            public string Source { get; private set; }
            public string Target { get; private set; }

            public Folders(string source, string target)
            {
                Source = source;
                Target = target;
            }
        }

        #endregion

        #region sadece aktif windowsdaki seçimleri yakala
        public static int NumberOfSelectedFiles; // seçili dosyaların sayısını tutacak

        public static List<string> FilesAndFolders(String ActiveWindowPath)
        {
            string filename;
            List<string> explorerItems = new List<string>();

            foreach (SHDocVw.InternetExplorer window in new SHDocVw.ShellWindows()) // Tüm Açık Klasörler i tarar
            {
                filename = Path.GetFileNameWithoutExtension(window.FullName).ToLower();
                
                if (filename.ToLowerInvariant() == "explorer")
                {
                   Shell32.FolderItems items = ((Shell32.IShellFolderViewDual2)window.Document).SelectedItems();
                   
                    foreach (Shell32.FolderItem item in items)
                    {
                        
                        if (ActiveWindowPath == Path.GetDirectoryName(item.Path))
                        {
                            explorerItems.Add(item.Path);
                            //string test = ((Shell32.IShellFolderViewDual2)window.Document).FocusedItem.Path;
                           // MessageBox.Show(item.Path);
                        }
                    }
                }
            }
            return explorerItems;
        }

        #endregion

        #region Windows API send key
        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        public static void PressKey(Keys key, bool up)
        {
            const int KEYEVENTF_EXTENDEDKEY = 0x1;
            const int KEYEVENTF_KEYUP = 0x2;
            if (up)
            {
                keybd_event((byte)key, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, (UIntPtr)0);
            }
            else
            {
                keybd_event((byte)key, 0x45, KEYEVENTF_EXTENDEDKEY, (UIntPtr)0);
            }
        }

        public static void F2KeySend()
        {
            //PressKey(Keys.ControlKey, false);
            PressKey(Keys.F2, false);
            PressKey(Keys.F2, true);
            //PressKey(Keys.ControlKey, true);
        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(String[] arguments)
        {
               
            #region multi instance control 
            // çoklu dosya seçiminde her seçilen dosya için bu uygulamadan tekrar çalıştırılır. fakat biz işlemlerimizi ilk çalışan uygulamada bitiriyoruz. 
            // uygulamanın tekrar çalıştığını algılayıp (yeni oluşturulan instancelar) bitirilir.
            bool createdNew;
            Mutex m = new Mutex(true, "myApp", out createdNew);

            if (!createdNew)
            {
                // MessageBox.Show("myApp is already running!", "Multiple Instances");

                return; // Uygulama zaten çalışıyor. O yüzden bunu kapat.  
            }

            #endregion

            #region Dosyaların Taşınması
            // Dosya/lar seçildi ise 
            if (arguments.Length > 0)
            {
                filePath = arguments[0];
                folderPath = Path.GetDirectoryName(arguments[0]) + "\\" + Path.GetFileNameWithoutExtension(arguments[0]);

                //- klasör .exe ile bitiyorsa windows oluşturmaya izin vermiyor. o yüzden .exe yi tespit edip .ese ile değiştiriyorum.
                string controlNameExe = Path.GetFileNameWithoutExtension(arguments[0]);
                if (controlNameExe.EndsWith(".exe"))
                {
                    string buff1 = controlNameExe.Remove(controlNameExe.Length - 4, 4) + ".ese";
                    folderPath = Path.GetDirectoryName(arguments[0]) + "\\" + buff1;
                }

                if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath); // Klasör Yoksa Oluştur

                NumberOfSelectedFiles = 0;

                // Seçili Dosyaların Tutuldugu string Koleksiyondan tek tek dosya yolunu al ve yeni oluşturulan klasöre taşı
                foreach (string e1 in FilesAndFolders(Path.GetDirectoryName(arguments[0])))
                {
                    //Gelen Adresin Türünü Belirleme File or Folder
                    FileAttributes attr = File.GetAttributes(e1);

                    if ((attr & FileAttributes.Directory) != FileAttributes.Directory)
                    {// File 
                        if (File.Exists(e1))
                        {
                            //MessageBox.Show("Kaynak: "  + e1 + "  ||  Hedef: " + folderPath + "\\" + Path.GetFileName(e1));
                            File.Move(e1, folderPath + "\\" + Path.GetFileName(e1));
                            NumberOfSelectedFiles++;
                        }
                    }
                    else
                    {// Folder
                        if (Directory.Exists(e1) && e1 != folderPath)
                        {
                            //MessageBox.Show("Kaynak: "  + e1 + "  ||  Hedef: " + folderPath);
                            MoveDirectory(e1, folderPath);
                            NumberOfSelectedFiles++;
                        }
                    }
                }

                // Klasör Oluşturulmuş ise Klasörü Seçili ve İsmini Değiştirmeye Hazır Hale Odaklan
                if (Directory.Exists(folderPath))
                {
                    ShowSelectedInExplorer.FilesOrFolders(folderPath); // Klasör Seçildi
                    Console.ReadLine();
                    // F2KeySend();
                }
            }
            else
            {
                // adres argumanı gelmedi ise:
                MessageBox.Show("Dosya Seçilmedi || Not Selected File");
            }

            #endregion
        }
    }

    /// <summary>
    /// string değişkelere isNumeric metodu ekle
    /// </summary>
    public static class ExtensionManager
    {
        public static bool IsNumeric(this string value)
        {
            double oReturn = 0;
            return double.TryParse(value, out oReturn);
        }
    }


}
