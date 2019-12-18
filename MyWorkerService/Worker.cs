/// https://tuhocict.com/thuc-thi-truy-van-sql-trong-c-lop-sqlcommand/

using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Spire.Pdf;
using Spire.Pdf.General.Find;
using Spire.Pdf.Graphics;
using Spire.Pdf.Security;
using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace MyWorkerService
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private string _connectionString = "Server=117.7.227.159;Database=ElectronicBill;User Id=sa;Password=Bach@khoa;MultipleActiveResultSets=true";

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        public override Task StartAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Service Started at: {time}", DateTimeOffset.Now);

            return base.StartAsync(cancellationToken);
        }

        public override Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Service Stopped at: {time}", DateTimeOffset.Now);
            return base.StopAsync(cancellationToken);
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            //while (!stoppingToken.IsCancellationRequested)
            //{


            //    await Task.Delay(1000, stoppingToken);
            //}


            QueryData();
        }

        private void QueryData()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                var command = new SqlCommand("SELECT * FROM Bills where ApproveStatus = 7", connection);
                var sqlDataReader = command.ExecuteReader();
                if (sqlDataReader.HasRows)
                {
                    while (sqlDataReader.Read())
                    {
                        var id = new Guid(sqlDataReader["Id"].ToString());
                        var unsignFileByte = sqlDataReader["UnsignFileByte"] as byte[];

                        System.IO.File.WriteAllBytes("abc.pdf", unsignFileByte);

                        // Load a PDF file and certificate
                        PdfDocument pdf = new PdfDocument();
                        pdf.LoadFromFile("abc.pdf");
                        int end = pdf.Pages.Count - 1;
                        PdfPageBase page = pdf.Pages[end];

                        PdfCertificate cert = new PdfCertificate("dtbk.pfx", "DienTuBachKhoa");

                        // Create a signature and set its position.
                        PdfSignature signature = new PdfSignature(pdf, page, cert, "Dinh Vi Bach Khoa");
                        signature.SignDetailsFont = new PdfTrueTypeFont(new Font("Times New Roman", 8f, FontStyle.Bold), true);
                        signature.SignFontColor = Color.Red;

                        signature.Bounds = new RectangleF(GetPositionSign(page), new SizeF(212, 40));
                        // Set the property for displaying the signature.
                        // Fill the contents of the signature.
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

                        // Set the document permission of the signature.
                        signature.DocumentPermissions = PdfCertificationFlags.ForbidChanges;

                        //Save pdf file.
                        string signedFileName = "signed" + ".pdf";
                        pdf.SaveToFile(signedFileName);
                        pdf.Close();

                        var updateQuery = "UPDATE [dbo].[Bills] SET [SignedFileByte] = @SignedFileByte WHERE [Id] = @Id";
                        //command.CommandText = updateQuery;
                        //command.Parameters.AddWithValue("SignedFileByte", lastName);
                        //count = command.ExecuteNonQuery();
                        _logger.LogInformation("Update Bill has Id = {0}", id.ToString());
                    }
                }

                sqlDataReader.Close();
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
    }
}
