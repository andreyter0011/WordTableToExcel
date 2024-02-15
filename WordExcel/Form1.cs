using Microsoft.Office.Interop.Word;
using excel = Microsoft.Office.Interop.Excel;
using word = Microsoft.Office.Interop.Word;

namespace WordExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void SelectWordFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word File|*.doc;*docx";
            openFileDialog.Title = "Âûבונטעו word פאיכ";
            if(openFileDialog.ShowDialog() ==DialogResult.OK)
            {
                ConvertWordToExcel(openFileDialog.FileName);              
            }
        }

        private void ConvertWordToExcel(string wordFilePath)
        {
            word.Application wordApp = new word.Application();
            word.Document wordDoc = null;
            try
            {
                wordDoc = wordApp.Documents.Open(wordFilePath);
                string tableText = "";
                foreach (word.Table table in wordDoc.Tables)
                {
                    foreach (word.Row row in table.Rows)
                    {
                        foreach (word.Cell cell in row.Cells)
                        {
                            tableText = cell.Range.Text + "\t";
                        }
                        tableText += "\n";
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}