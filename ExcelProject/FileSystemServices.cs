using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelProject
{
    class FileSystemServices
    {
        private static string[] subdirectories = new string[] 
        {
            "01 Kontakte","02 Gesprächsnotizen","03 SVK", "04 Berichte", 
            "05 Angebot, Vertrag und Rechnungen","06 Berechnungen, Ausarbeitungen",
            "07 Literatur","08 Fotodokumentation","09 Messdaten (Labor)","10 Daten von Dritten (Zuarbeit)", "11 Veröffentlichungen"
        };
        public static string ProjectsRootFolderTreeDir = "D:\\Test\\";

        /// <summary>
        /// create directory of a given project with its subdirectories
        /// </summary>
        /// <param name="projectNumber"></param>
        /// <param name="projectName"></param>
        public static void CreateFolders(string projectNumber, string projectName)
        {
            //string fullDirectoryName = projectNumber + "_" + projectName;
            string fullPathName = ProjectsRootFolderTreeDir + projectNumber + "_" + projectName + "\\";
                foreach (string subdir in FileSystemServices.subdirectories)
                {
                    Directory.CreateDirectory(fullPathName + subdir);
                }
            
        }

        public static void OpenProjectFolder(string projectNumber, string projectName)
        {
            string fullPathName = ProjectsRootFolderTreeDir + projectNumber + "_" + projectName;
            string fullPathNameToOpen = ProjectsRootFolderTreeDir + projectNumber + "_" + projectName + "\\";
            if (!Directory.Exists(fullPathName))
            {
                System.Windows.Forms.MessageBox.Show("Verzeichniss:\n"+fullPathName + "\nexistiert nicht!");
                return;
            }

            System.Diagnostics.Process.Start("explorer.exe",fullPathNameToOpen);
        }
    }
}
