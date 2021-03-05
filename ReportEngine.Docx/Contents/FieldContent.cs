using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace ReportEngine.Docx.Contents
{
    public class FieldContent
    {
        public FieldContent()
        {
        }

        public FieldContent(string name)
        {
            Name = name;
        }

        public FieldContent(string name, object value)
            : this(name)
        {
            Value = value;
        }

        public FieldContent(string name, Image image)
            : this(name)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                image.Save(stream, ImageFormat.Jpeg);
                Value = stream.ToArray();
            }
        }

        public FieldContent(string name, byte[] image)
            : this(name)
        {
            Value = image;
        }

        public FieldContent(string name, string imagePathOrValue, bool isImage)
            : this(name)
        {
            if (!isImage)
                Value = imagePathOrValue;
            else
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(imagePathOrValue);

                    Value = new byte[fileInfo.Length];
                    using (FileStream fs = fileInfo.OpenRead())
                        fs.Read(((byte[])Value), 0, ((byte[])Value).Length);
                }
                catch
                {
                    Value = imagePathOrValue;
                }
            }
        }

        public string Name { get; set; }
        public object Value { get; set; }
    }
}
