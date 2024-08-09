using System.Reflection.Metadata;
using System.Windows;
using System.Windows.Documents;
using Microsoft.Office.Interop.Word;
using Document = Microsoft.Office.Interop.Word.Document;
using Application = Microsoft.Office.Interop.Word.Application;
using Window = System.Windows.Window;
using Paragraph = Microsoft.Office.Interop.Word.Paragraph;
using Microsoft.Win32;
using Range = Microsoft.Office.Interop.Word.Range;
using System.Security.Cryptography.Xml;

namespace WpfApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BrowseTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Word Documents|*.docx",
                Title = "Chọn đường dẫn để lấy template"
            };

            if (dialog.ShowDialog() == true)
            {
                TemplatePathTextBox.Text = dialog.FileName;
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Chọn đường dẫn để lấy template"
            };

            if (dialog.ShowDialog() == true)
            {
                SavePathTextBox.Text = dialog.FileName;
            }
        }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
            string templatePath = TemplatePathTextBox.Text;
            string savePath = SavePathTextBox.Text;

            if (string.IsNullOrEmpty(templatePath) || string.IsNullOrEmpty(savePath))
            {
                MessageBox.Show("Vui lòng chọn đường dẫn template và lưu file.");
                return;
            }

            if (!IsValidEmail(TeacherEmailTextBox.Text))
            {
                MessageBox.Show("Email giáo viên không hợp lệ");
                return;
            }

            if (LeaveDatePicker.SelectedDate == null)
            {
                MessageBox.Show("Vui lòng chọn ngày nghỉ");
                return;
            }

            if (CurrentDatePicker.SelectedDate == null)
            {
                MessageBox.Show("Vui lòng chọn ngày hiện tại");
                return;
            }

            var replacements = new Dictionary<string, string>
                {
                    { "(FullName)", FullNameTextBox.Text },
                    { "(ClassCode)", ClassCodeTextBox.Text },
                    { "(Subject)", SubjectTextBox.Text },
                    { "(LeaveDate)", LeaveDatePicker.SelectedDate?.ToShortDateString() ?? string.Empty },
                    { "(Reason)", ReasonTextBox.Text },
                    { "(TeacherName)", TeacherNameTextBox.Text },
                    { "(TeacherEmail)", TeacherEmailTextBox.Text },
                    { "(CurrentDate)", CurrentDatePicker.SelectedDate?.ToShortDateString() ?? string.Empty },
                    { "(Signature)", SignatureTextBox.Text }
                };

            SaveToWord(templatePath, savePath, replacements);
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        private void SaveToWord(string templatePath, string savePath, Dictionary<string, string> replacements)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = null;

            try
            {
                doc = wordApp.Documents.Open(templatePath);
                foreach (var entry in replacements)
                {
                    FindAndReplace(wordApp, entry.Key, entry.Value);
                }

                doc.SaveAs2(savePath);
                string pdfFilePath = System.IO.Path.ChangeExtension(savePath, ".pdf");
                doc.ExportAsFixedFormat(pdfFilePath, WdExportFormat.wdExportFormatPDF);

                MessageBox.Show("File đã được lưu thành công tại: " + savePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                }
                wordApp.Quit();
            }
        }

        private void FindAndReplace(Application wordApp, string findText, string replaceText)
        {
            foreach (Range range in wordApp.ActiveDocument.StoryRanges)
            {
                range.Find.ClearFormatting();
                range.Find.Text = findText;
                range.Find.Replacement.ClearFormatting();
                range.Find.Replacement.Text = replaceText;

                range.Find.Execute(Replace: WdReplace.wdReplaceAll);
            }
        }

    }
}
