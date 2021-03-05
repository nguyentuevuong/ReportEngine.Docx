using System.Collections.Generic;

namespace ReportEngine.Docx.Contents
{
    public class TableContent
    {
        public TableContent()
        {

        }

        public TableContent(string name)
        {
            Name = name;
        }

        public TableContent(string name, IEnumerable<IEnumerable<FieldContent>> rows)
            : this(name)
        {
            Rows = rows;
        }

        public TableContent(string name, params IEnumerable<FieldContent>[] rows)
            : this(name)
        {
            Rows = rows;
        }

        public string Name { get; set; }
        public IEnumerable<IEnumerable<FieldContent>> Rows { get; set; }
    }
}
