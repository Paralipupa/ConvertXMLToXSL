using System.Windows;



namespace Egrn.Model
{
    public class DialogService
    {
        public string File { get; set; }
        public string Path { get; set; }

        public  bool OpenFileDialog(string filter = "xml(*.xml)|*.xml|All files (*.*)|*.*")
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();

            openFileDialog.Title = "Выберите файл";
            if (string.IsNullOrWhiteSpace(File) == false)
            {
                openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(File);
            }
            openFileDialog.Filter = filter;
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == true)
            {
                File = openFileDialog.FileName;
                return true;
            }
            return false;
        }

        public bool OpenPathDialog()
        {            
            
            System.Windows.Forms.FolderBrowserDialog openPathDialog = new System.Windows.Forms.FolderBrowserDialog();
            if (string.IsNullOrWhiteSpace(Path))
            {
                Path = System.AppDomain.CurrentDomain.BaseDirectory;
            }
            openPathDialog.SelectedPath = Path;

            if (openPathDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Path = openPathDialog.SelectedPath;
                return true;
            }
            return false;
        }


        public static void ShowMessage(string message)
        {
            MessageBox.Show(message, "", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public static MessageBoxResult MsgBox(string message)
        {
            return MessageBox.Show(message, "Внимание", MessageBoxButton.YesNo);
        }
    }

}
