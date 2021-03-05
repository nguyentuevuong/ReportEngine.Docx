using System.Collections.Generic;

namespace ReportEngine.Docx.Contents
{
    public class Content
    {
        public IEnumerable<TableContent> Tables { get; set; }
        public IEnumerable<FieldContent> Fields { get; set; }
    }
}
