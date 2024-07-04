using System;
using System.IO;
using System.Reflection.Metadata;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;
using Outlook = Microsoft.Office.Interop.Outlook;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Xml.Linq;
using System.Xml;
using System.Text.RegularExpressions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using Microsoft.Office.Interop.Word;
namespace OutlookExcelEmail
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
            txtEmailSubject.Text = @"����IPS H8.5G LD PJT �۾�����(HIT)_" + DateTime.Now.ToString("yyyyMMdd") + " �ۺΰ�";
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            string excelFilePath = txtExcelFilePath.Text;
            string emailTo = txtTestEmail.Text;
            string emailSubject = txtEmailSubject.Text;
            
            string emailBody = GetExcelContent(excelFilePath);
            if(emailBody == "")
            {
                //MessageBox.Show("���� ������ ���� �ҷ����µ� �����߽��ϴ�!. �ٽ� �õ����ּ���.");
                return;
            }

            SendEmailWithAttachment(excelFilePath, emailTo, "", emailSubject, emailBody);
            MessageBox.Show("Test email sent successfully.");
            
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            try
            {
                string excelFilePath = txtExcelFilePath.Text;
                string emailTo = txtSendEmail.Text;
                string emailSubject = txtEmailSubject.Text;
                string emailBody = GetExcelContent(excelFilePath);

                if (emailBody == "")
                {
                    //MessageBox.Show("���� ������ ���� �ҷ����µ� �����߽��ϴ�!. �ٽ� �õ����ּ���.");
                    return;
                }

                SendEmailWithAttachment(excelFilePath, emailTo, "", emailSubject, emailBody);
                MessageBox.Show("Email sent successfully.");
            }
            catch
            {

            }
            
        }

        private string GetExcelContent(string excelFilePath)
        {
            try
            {
                Application excelApp = new Application();
                Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
                Worksheet worksheet = workbook.Worksheets.get_Item(workbook.Worksheets.Count) as Worksheet;

                Range range = worksheet.UsedRange;
                object[,] data = (object[,])range.Value;
                
                int Readrows = 7;
                int Readcolumns = 1;
                int Startrows = 11;
                int Startcolumns = 1;
                string content = "";

                for (int i = Startrows; i <= Startrows+Readrows; i++)
                {
                    for (int j = Startcolumns; j <= Startcolumns+Readcolumns; j++)
                    {
                        if (data[i, j] == null)
                        {
                            continue;
                        }
                        else if (data[i, j] is string)
                        {

                            content += "<tr>";
                            if (i == Startrows)
                            {
                                content += "<td colspan=\"44\" data-editing-info=\"{&quot;borderOverride&quot;:true}\" style=\"text-align: left; border-top: 1.333px solid rgb(0, 0, 0); border-right: 1pt solid black; border-left: 1.333px solid rgb(0, 0, 0); padding-top: 1px; padding-right: 1px; padding-left: 1px; vertical-align: middle; width: 644pt; height: 16.5pt;\">";
                            }
                            else
                                content += "<td colspan=\"44\" data-editing-info=\"{&quot;borderOverride&quot;:true}\" style=\"text-align: left; border-right: 1pt solid black; border-left: 1.333px solid rgb(0, 0, 0); padding-top: 1px; padding-right: 1px; padding-left: 1px; vertical-align: middle; width: 644pt; height: 16.5pt;\">\r\n";


                            content += data[i, j] as string;
                            content += "</td>";
                            content += "</tr>";
                        }
                        else
                        {
                            content += data[i, j].ToString();
                        }

                        //content += data[i, j].ToString() + "\t";
                    }
                    content += "<br>";
                }
                content += "<td colspan=\"44\" data-editing-info=\"{&quot;borderOverride&quot;:true}\" style=\"text-align: left; border-right: 1pt solid black; border-bottom: 1.333px solid rgb(0, 0, 0); border-left: 1.333px solid rgb(0, 0, 0); padding-top: 1px; padding-right: 1px; padding-left: 1px; vertical-align: middle; width: 644pt; height: 16.5pt;\">";
                workbook.Close(false);
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                return content;
            }
            catch(System.Exception e)
            {
                MessageBox.Show("���� ������ �ҷ����µ� �����߽��ϴ�! �ٽ� �õ����ּ���. \nError : " +
                    e.ToString());
                return "";
            }
            return "";


        }

        private void SendEmailWithAttachment(string attachmentPath, string to, string cc, string subject, string body)
        {

            string Sign1, Sign2;

            Outlook.Application outlookApp = new Outlook.Application();
            MailItem mailItem = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);
            Inspector inspector = mailItem.GetInspector;
            var sig = inspector.WordEditor as Microsoft.Office.Interop.Word.Document;
            

            mailItem.Subject = subject;
            mailItem.To = to;
            if(cc != "")
            {
                mailItem.CC = cc;
            }
            mailItem.Display(mailItem);
            //Sign2;
            GetSignatureParts(mailItem.HTMLBody, out Sign1, out Sign2);

            string html = $"{Sign1}" + "<br><br>";//�ȳ��ϼ���~ �� ù ��° ����
            html += $"8.5G LD PJT {DateTime.Now.ToString("yyyyMMdd")} ���������� �ۺε帳�ϴ�." + "<br><br>";//�������� �ۺ�
            html += "<table id=\"x_table_0\" style=\"width: 644pt; box-sizing: border-box; border-collapse: collapse; border-spacing: 0px;\">"; //ǥ �׸���
            html += "<tbody>"; //ǥ �׸���
            html += body; //���� ����
            html += "</tbody>";
            html += "</table><br><br>";
            html += "�ڼ��� ������ ÷�������� Ȯ���� �ּ���." + "<br><br>"; //÷������ Ȯ�� �޼���

            html += $"{Sign2}"; //������ ����
            mailItem.HTMLBody = html;
                
                
                
                
                
                 

            if (File.Exists(attachmentPath))
            {
                mailItem.Attachments.Add(attachmentPath, OlAttachmentType.olByValue, Type.Missing, Type.Missing);
            }
            else
            {
                MessageBox.Show("The attachment file does not exist.");
            }

            mailItem.Send();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
        }

        private void GetSignatureParts(string body, out string part1, out string part2)
        {
            // Get the default signature
            string fullSignature = body;

            // Split the signature into two parts
            // For the sake of the example, let's assume the first line is the first part
            // and the rest is the second part
            var lines = fullSignature.Split(new[] { "</a>" }, StringSplitOptions.None);

            part1 = lines[0] + "</a>";
            part2 = string.Join("<br>", lines, 1, lines.Length - 1);
        }

    }
}