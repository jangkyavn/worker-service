using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Spire.Pdf;
using Spire.Pdf.General.Find;
using Spire.Pdf.Graphics;
using Spire.Pdf.Security;
using System;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace MyWorkerService
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private string _baseUrl = "https://localhost:44383/";
        // private string _baseUrl = "http://hoadon.dvbk.vn/";
        private string _connectionString = "Server=117.7.227.159;Database=ElectronicBill;User Id=sa;Password=Bach@khoa;MultipleActiveResultSets=true";
        // private string _connectionString = "Server=103.63.109.19;Database=VanLanElectronicBill;User Id=sa;Password=PhanMem@BachKhoa;MultipleActiveResultSets=true";

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        // Bắt đầu chạy service
        public override Task StartAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Service Started at: {time}", DateTimeOffset.Now);

            return base.StartAsync(cancellationToken);
        }

        // Dừng chạy service
        public override Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Service Stopped at: {time}", DateTimeOffset.Now);
            return base.StopAsync(cancellationToken);
        }

        // Thực thi sau bắt đầu chạy service
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            // Vòng lặp vô hạn
            while (!stoppingToken.IsCancellationRequested)
            {
                // Ký hóa đơn
                await SignBillAsync();
                // Ký biên bản
                await SignRecordAsync();

                // Timer cứ 1s là thực hiện hành động
                await Task.Delay(1000, stoppingToken);
            }
        }

        private async Task SignBillAsync()
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                SqlCommand billCommand = new SqlCommand("SELECT * FROM Bills where ApproveStatus = 7", connection);
                SqlDataReader billReader = await billCommand.ExecuteReaderAsync();
                _logger.LogInformation("HasRows: {0}", billReader.HasRows);
                if (billReader.HasRows)
                {
                    if (billReader.Read())
                    {
                        Guid id = new Guid(billReader["Id"].ToString());

                        SqlCommand billByteCommand = new SqlCommand($"SELECT * FROM BillBytes where BillId = '{id}'", connection);
                        SqlDataReader billByteReader = await billByteCommand.ExecuteReaderAsync();

                        if (billByteReader.HasRows)
                        {
                            if (billByteReader.Read())
                            {
                                byte[] unsignFileByte = billByteReader["UnsignPdfByte"] as byte[];
                                string unsignFileName = id + "unsign.pdf";

                                await File.WriteAllBytesAsync(unsignFileName, unsignFileByte);

                                // Load a PDF file and certificate
                                PdfDocument pdf = new PdfDocument();
                                pdf.LoadFromFile(unsignFileName);
                                int end = pdf.Pages.Count - 1;
                                PdfPageBase page = pdf.Pages[end];

                                PdfCertificate cert = new PdfCertificate("dtbk.pfx", "DienTuBachKhoa");

                                // Create a signature and set its position.
                                PdfSignature signature = new PdfSignature(pdf, page, cert, "Dinh Vi Bach Khoa");
                                signature.SignDetailsFont = new PdfTrueTypeFont(new Font("Times New Roman", 8f, FontStyle.Bold), true);
                                signature.SignFontColor = Color.Red;

                                signature.Bounds = new RectangleF(GetPositionSign(page), new SizeF(212, 40));
                                signature.NameLabel = "Signature Valid\nKý bởi: ";
                                signature.Name = string.Format("CT CP THIẾT BỊ ĐIỆN - ĐIỆN TỬ BÁCH KHOA\nNgày ký: {0}\n\n\n", DateTime.Now.ToString("dd/MM/yyyy"));
                                signature.DistinguishedName = "ĐỊNH VỊ BÁCH KHOA";

                                signature.ReasonLabel = "\nReason: ";
                                signature.Reason = "Hóa đơn giá trị gia tăng";                                                      // Add

                                signature.DateLabel = "\nNgày ký: ";
                                signature.Date = DateTime.Now;

                                signature.ContactInfoLabel = "Tổng đài CSKH: ";
                                signature.ContactInfo = "1900 5555 13";                                                             // Add
                                signature.LocationInfo = "Số 561 Nguyễn Bỉnh Khiêm, Hải An, Hải Phòng, Việt Nam";                   // Aad

                                signature.Certificated = false;

                                signature.DocumentPermissions = PdfCertificationFlags.ForbidChanges;

                                string signedFileName = id + "signed" + ".pdf";
                                pdf.SaveToFile(signedFileName);
                                pdf.Close();

                                SqlCommand updateBillCommand = new SqlCommand("UPDATE [dbo].[BillBytes] SET [SignedPdfByte] = @SignedPdfByte WHERE [BillId] = @BillId", connection);
                                updateBillCommand.Parameters.AddWithValue("BillId", id);
                                updateBillCommand.Parameters.AddWithValue("SignedPdfByte", File.ReadAllBytes(signedFileName));
                                await updateBillCommand.ExecuteNonQueryAsync();

                                if (!string.IsNullOrEmpty(unsignFileName))
                                {
                                    if (File.Exists(unsignFileName))
                                    {
                                        File.Delete(unsignFileName);
                                    }
                                }

                                if (!string.IsNullOrEmpty(signedFileName))
                                {
                                    if (File.Exists(signedFileName))
                                    {
                                        File.Delete(signedFileName);
                                    }
                                }

                                using (HttpClient client = new HttpClient())
                                {
                                    client.BaseAddress = new Uri(_baseUrl);
                                    client.DefaultRequestHeaders.Accept.Clear();
                                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                                    try
                                    {
                                        await client.PostAsJsonAsync("api/Bill/postSignedBill", new
                                        {
                                            Id = id
                                        });
                                    }
                                    catch (Exception ex)
                                    {
                                        _logger.LogError("Error Post Data: {0}", ex);
                                    }
                                }

                                await billByteReader.CloseAsync();

                                _logger.LogInformation("Update Bill has Id = {0}", id.ToString());
                            }
                        }
                    }
                }

                await billReader.CloseAsync();
            }
        }

        private async Task SignRecordAsync()
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                SqlCommand recordCommand = new SqlCommand("SELECT * FROM Records where ApproveStatus = 7", connection);
                SqlDataReader recordReader = await recordCommand.ExecuteReaderAsync();
                if (recordReader.HasRows)
                {
                    if (recordReader.Read())
                    {
                        Guid id = new Guid(recordReader["Id"].ToString());

                        SqlCommand recordByteCommand = new SqlCommand($"SELECT * FROM RecordBytes where RecordId = '{id}'", connection);
                        SqlDataReader recordByteReader = await recordByteCommand.ExecuteReaderAsync();

                        if (recordByteReader.HasRows)
                        {
                            if (recordByteReader.Read())
                            {
                                byte[] unsignFileByte = recordByteReader["UnsignPdfByte"] as byte[];
                                string unsignFileName = id + "unsign.pdf";

                                await File.WriteAllBytesAsync(unsignFileName, unsignFileByte);

                                // Load a PDF file and certificate
                                PdfDocument pdf = new PdfDocument();
                                pdf.LoadFromFile(unsignFileName);
                                int end = pdf.Pages.Count - 1;
                                PdfPageBase page = pdf.Pages[end];

                                PdfCertificate cert = new PdfCertificate("dtbk.pfx", "DienTuBachKhoa");

                                // Create a signature and set its position.
                                PdfSignature signature = new PdfSignature(pdf, page, cert, "Dinh Vi Bach Khoa");
                                signature.SignDetailsFont = new PdfTrueTypeFont(new Font("Times New Roman", 8f, FontStyle.Bold), true);
                                signature.SignFontColor = Color.Red;

                                signature.Bounds = new RectangleF(GetPositionSignRecored(page), new SizeF(212, 40));
                                signature.NameLabel = "Signature Valid\nKý bởi: ";
                                signature.Name = string.Format("CT CP THIẾT BỊ ĐIỆN - ĐIỆN TỬ BÁCH KHOA\nNgày ký: {0}\n\n\n", DateTime.Now.ToString("dd/MM/yyyy"));
                                signature.DistinguishedName = "ĐỊNH VỊ BÁCH KHOA";

                                signature.ReasonLabel = "\nReason: ";
                                signature.Reason = "Hóa đơn giá trị gia tăng";                                                      // Add

                                signature.DateLabel = "\nNgày ký: ";
                                signature.Date = DateTime.Now;

                                signature.ContactInfoLabel = "Tổng đài CSKH: ";
                                signature.ContactInfo = "1900 5555 13";                                                             // Add
                                signature.LocationInfo = "Số 561 Nguyễn Bỉnh Khiêm, Hải An, Hải Phòng, Việt Nam";                   // Aad

                                signature.Certificated = false;

                                signature.DocumentPermissions = PdfCertificationFlags.ForbidChanges;

                                string signedFileName = id + "signed" + ".pdf";
                                pdf.SaveToFile(signedFileName);
                                pdf.Close();

                                SqlCommand updateRecordCommand = new SqlCommand("UPDATE [dbo].[RecordBytes] SET [SignedPdfByte] = @SignedPdfByte WHERE [RecordId] = @RecordId", connection);
                                updateRecordCommand.Parameters.AddWithValue("RecordId", id);
                                updateRecordCommand.Parameters.AddWithValue("SignedPdfByte", File.ReadAllBytes(signedFileName));
                                await updateRecordCommand.ExecuteNonQueryAsync();

                                if (!string.IsNullOrEmpty(unsignFileName))
                                {
                                    if (File.Exists(unsignFileName))
                                    {
                                        File.Delete(unsignFileName);
                                    }
                                }

                                if (!string.IsNullOrEmpty(signedFileName))
                                {
                                    if (File.Exists(signedFileName))
                                    {
                                        File.Delete(signedFileName);
                                    }
                                }

                                using (HttpClient client = new HttpClient())
                                {
                                    client.BaseAddress = new Uri(_baseUrl);
                                    client.DefaultRequestHeaders.Accept.Clear();
                                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                                    try
                                    {
                                        await client.PostAsJsonAsync("api/Record/postSignedRecord", new
                                        {
                                            Id = id
                                        });
                                    }
                                    catch (Exception ex)
                                    {
                                        _logger.LogError("Error Post Data: {0}", ex);
                                    }
                                }

                                _logger.LogInformation("Update Record has Id = {0}", id.ToString());
                            }
                        }

                        await recordByteReader.CloseAsync();
                    }
                }

                await recordReader.CloseAsync();
            }
        }

        private PointF GetPositionSign(PdfPageBase page)
        {
            // Create default
            PointF point = new PointF(335, 685);

            PdfTextFind[] result = page.FindText("Người bán hàng (Seller)").Finds;

            if (result != null && result.Count() > 0)
            {
                point = result[0].Position;

                point.X -= 35;
                point.Y += 20;
            }

            return point;
        }

        private PointF GetPositionSignRecored(PdfPageBase page)
        {
            // Create default
            PointF point = new PointF(335, 685);

            PdfTextFind[] result = page.FindText("ĐẠI DIỆN BÊN B").Finds;

            if (result != null && result.Count() > 0)
            {
                point = result[0].Position;

                point.X -= 35;
                point.Y += 20;
            }

            return point;
        }
    }
}
