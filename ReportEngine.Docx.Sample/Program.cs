using System;
using System.Collections.Generic;
using System.Linq;
using ReportEngine.Docx;
using ReportEngine.Docx.Contents;
using System.IO;

namespace ReportEngine.Docx.Sample
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] arg)
        {
            var value2Fill = new Content
            {
                Fields = new List<FieldContent>
                {
                    new FieldContent("Địa điểm","Hà Nội"),
                    new FieldContent("Ngày","17"),
                    new FieldContent("Tháng","02"),
                    new FieldContent("Năm","2016"),
                    new FieldContent("Tên nhiệm vụ", "Nhiệm vụ KH&CN 1"),
                    new FieldContent("Mã số","HJJ5345"),
                    new FieldContent("Tổ chức chủ trì","Tổ chức 1"),
                    new FieldContent("Chủ nhiệm","Lê Văn C"),
                    new FieldContent("Họ tên chuyên gia","Nguyễn Văn A"),
                    new FieldContent("Ngày nhận","23"),
                    new FieldContent("Tháng nhận","04"),
                    new FieldContent("Năm nhận","2016"),
                    new FieldContent("Ý kiến khác","Ý kiến...")
                },

                Tables = new List<TableContent>
                {
                    new TableContent
                    {
                        Name = "Bảng đánh giá SL KL",
                        Rows = new List<List<FieldContent>>()
                    },

                    new TableContent
                    {
                        Name = "Bảng chất lượng",
                        Rows = new List<List<FieldContent>>()
                    },
                }
            };
            var rows1 = new List<List<FieldContent>>();
            for (int i = 1; i < 6; i++)
                rows1.Add(new List<FieldContent>()
            {
                new FieldContent { Name = "STT", Value = i  },
                new FieldContent { Name = "Tên sản phẩm", Value ="Sản phẩm..." },
                new FieldContent { Name = "Đặt hàng", Value  = "100" },
                new FieldContent { Name = "Đạt được", Value = "99" },
                new FieldContent { Name = "Đạt", Value = "X" },
                new FieldContent { Name = "Không đạt", Value = "X" },
                new FieldContent { Name = "Ghi chú", Value = "..." }
            });
            value2Fill.Tables.FirstOrDefault(table => table.Name == "Bảng đánh giá SL KL").Rows = rows1;

            var rows2 = new List<List<FieldContent>>();
            for (int i = 1; i < 6; i++)
                rows2.Add(new List<FieldContent>()
            {
                new FieldContent { Name = "STT", Value = i  },
                new FieldContent { Name = "Tên sản phẩm", Value ="Sản phẩm..." },
                new FieldContent { Name = "Đặt hàng", Value  = "100" },
                new FieldContent { Name = "Đạt được", Value = "99" },
                new FieldContent { Name = "Đạt", Value = "X" },
                new FieldContent { Name = "Không đạt", Value = "X" },
                new FieldContent { Name = "Ghi chú", Value = "..." }
            });
            value2Fill.Tables.FirstOrDefault(table => table.Name == "Bảng chất lượng").Rows = rows2;

            using (var outputDocument = new DocxTemplate("PL5-PĐGKQ.docx"))
                outputDocument.FillContent(value2Fill);
        }
    }
}
