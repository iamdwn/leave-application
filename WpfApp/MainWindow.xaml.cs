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
using System.Net.Mail;
using System.Net;
using Aspose.Words;
using System.Windows.Controls;
using System.IO;
using System.Text;

namespace WpfApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        string pdfFilePath;

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
                Title = "Chọn đường dẫn để lưu"
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
                pdfFilePath = System.IO.Path.ChangeExtension(savePath, ".pdf");
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

        private void SendMailButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SendEmail(TemplatePathTextBox.Text, pdfFilePath);
                MessageBox.Show("Gửi email thành công!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Gửi email thất bại!");
            }
        }

        private async void SendEmail(string attachment1, string attachment2)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = null;

            try
            {
                doc = wordApp.Documents.Open(TemplatePathTextBox.Text);

                string htmlFilePath = System.IO.Path.ChangeExtension(TemplatePathTextBox.Text, ".html");
                doc.SaveAs2(htmlFilePath, WdSaveFormat.wdFormatHTML);

                if (doc != null)
                {
                    doc.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                    doc = null;
                }

                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                wordApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                string htmlBody = File.ReadAllText(htmlFilePath, Encoding.UTF8);
                string subject = $"Đơn xin phép nghỉ học";
                //string body = @"<html><head><meta charset=""UTF-8""></head><body>" + htmlBody + "</body></html>";
                string body = $@"
                            <p>ĐƠN XIN NGHỈ PHÉP</p>
                            <p><strong>Kính gửi:</strong> {TeacherNameTextBox.Text}</p>
                            <p><strong>Tên tôi là:</strong> {FullNameTextBox.Text}</p>
                            <p><strong>Mã lớp:</strong> {ClassCodeTextBox.Text}</p>
                            <p><strong>Tên môn:</strong> {SubjectTextBox.Text}</p>
                            <p><strong>Ngày nghỉ:</strong> {LeaveDatePicker.SelectedDate?.ToShortDateString() ?? string.Empty}</p>
                            <p><strong>Lí do nghỉ:</strong> {ReasonTextBox.Text}</p>
                            <p><strong>Email giáo viên:</strong> {TeacherEmailTextBox.Text}</p>
                            <p>Tôi xin phép được nghỉ học. Tôi cam kết sẽ bổ sung và hoàn thành bài tập và kiến thức đã lỡ sau khi trở lại lớp.</p>
                            <p>Kính mong giáo viên chấp thuận.</p>
                            <p>Xin chân thành cảm ơn!</p>
                            <p><strong>Ngày:</strong> {CurrentDatePicker.SelectedDate?.ToShortDateString() ?? string.Empty}</p>
                            <p><strong>Ký tên:</strong> {SignatureTextBox.Text}</p>";

                System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage()
                {
                    From = new MailAddress("nextintern.corp@gmail.com", "Duong Truong"),
                    Subject = subject,
                    IsBodyHtml = true,
                    BodyEncoding = Encoding.UTF8,
                    SubjectEncoding = Encoding.UTF8
                };

                mail.Attachments.Add(new System.Net.Mail.Attachment(attachment1));
                mail.Attachments.Add(new System.Net.Mail.Attachment(attachment2));
                AlternateView alternateView = AlternateView.CreateAlternateViewFromString(body, Encoding.UTF8, "text/html");
                mail.AlternateViews.Add(alternateView);

                mail.To.Add(TeacherEmailTextBox.Text);

                SmtpClient smtpClient = new SmtpClient("smtp.gmail.com")
                {
                    Port = 587,
                    Credentials = new NetworkCredential("nextintern.corp@gmail.com", "wflm cyhu ifww lnbz"),
                    EnableSsl = true
                };

                await smtpClient.SendMailAsync(mail);
            }
            catch (Exception ex)
            {

                MessageBox.Show($"Error sending email to {TeacherNameTextBox.Text} ({TeacherEmailTextBox.Text}): {ex.Message}");
            }
            finally
            {
                if (doc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }
                if (wordApp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
            }
        }

    }
}
