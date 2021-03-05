using ReportEngine.Docx.Contents;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using d = DocumentFormat.OpenXml.Drawing;
using o = DocumentFormat.OpenXml;
using p = DocumentFormat.OpenXml.Packaging;
using w = DocumentFormat.OpenXml.Wordprocessing;

namespace ReportEngine.Docx
{
    public class DocxTemplate : IDisposable
    {
        private class Embed
        {
            public string Name { get; set; }
            public byte[] Image { get; set; }
        }

        private List<Embed> Embeds = new List<Embed>();
        private string fileInput = "", fileOutput = "";
        private p.WordprocessingDocument wordDocument;

        #region constructor
        public DocxTemplate(string fileInput)
        {
            this.fileInput = fileInput;
            this.fileOutput = CreateFileOutput(fileInput, null);
            wordDocument = p.WordprocessingDocument.Open(fileOutput, true);
        }

        public DocxTemplate(string fileInput, string fileOutput)
        {
            this.fileInput = fileInput;
            this.fileOutput = fileOutput;
            try
            {
                File.Delete(fileOutput);
                File.Copy(fileInput, fileOutput);
            }
            catch
            {
                this.fileOutput = CreateFileOutput(this.fileInput, 0);
            }
            wordDocument = p.WordprocessingDocument.Open(this.fileOutput, true);
        }
        #endregion

        #region publish method
        public void FillContent(Content content)
        {
            if (content != null && !fileOutput.Equals("NULL"))
            {
                if (content.Fields != null && content.Fields.Count() > 0)
                    foreach (var field in content.Fields)
                        if (field != null && !String.IsNullOrEmpty(field.Name))
                        {
                            w.SdtElement sdtElement = wordDocument.MainDocumentPart.Document.Body.Descendants<w.SdtElement>()
                                .FirstOrDefault(c =>
                                {
                                    w.SdtProperties properties = c.Elements<w.SdtProperties>().FirstOrDefault();
                                    if (properties != null)
                                    {
                                        w.SdtAlias alias = properties.Elements<w.SdtAlias>().FirstOrDefault();
                                        if (alias != null && alias.Val.HasValue && alias.Val.Value.Trim() == field.Name.Trim())
                                            return true;
                                    }
                                    return false;
                                });

                            if (sdtElement != null)
                                FillField(field, sdtElement);
                        }

                if (content.Tables != null && content.Tables.Count() > 0)
                    foreach (var table in content.Tables)
                        if (table != null && table.Rows != null && table.Rows.Count() > 0)
                        {
                            w.SdtElement sdtElement = wordDocument.MainDocumentPart.Document.Body.Descendants<w.SdtElement>()
                                .FirstOrDefault(c =>
                                {
                                    w.SdtProperties properties = c.Elements<w.SdtProperties>().FirstOrDefault();
                                    if (properties != null)
                                    {
                                        w.SdtAlias alias = properties.Elements<w.SdtAlias>().FirstOrDefault();
                                        if (alias != null && alias.Val.HasValue && alias.Val.Value.Trim() == table.Name.Trim())
                                            return true;
                                    }
                                    return false;
                                });

                            if (sdtElement != null)
                                FillTable(table, sdtElement);
                        }
            }

            var checkeed = true;
            while (checkeed)
            {
                w.SdtElement sdtElem = wordDocument.MainDocumentPart.Document.Body.Descendants<w.SdtElement>().FirstOrDefault();
                if (sdtElem != null)
                {
                    o.OpenXmlElement newelement;
                    if (sdtElem as w.SdtRun != null) //run
                        newelement = sdtElem.LastChild.CloneNode(true);
                    else // block
                        newelement = sdtElem.LastChild.LastChild.CloneNode(true);

                    if (newelement as w.BookmarkEnd != null)
                        newelement = new w.Paragraph(sdtElem.Descendants<w.Run>().FirstOrDefault().CloneNode(true));

                    sdtElem.Parent.InsertAfter(newelement, sdtElem);
                    sdtElem.Parent.RemoveChild(sdtElem);
                }
                else
                    checkeed = false;
            }
        } 
        #endregion

        #region private method
        private void FillField(FieldContent field, w.SdtElement sdtElement)
        {
            if (field.Value as byte[] == null || ((byte[])field.Value).Length == 0)
            {
                w.Text text4edit =
                    sdtElement.Descendants<w.Text>()
                        .FirstOrDefault();
                if (text4edit != null)
                    text4edit.Text = field.Value != null ? field.Value.ToString() : "";
                else
                    sdtElement.Append(
                        new w.SdtContentBlock(new w.Paragraph(
                            new w.Run(
                                new w.Text(field.Value != null ? field.Value.ToString() : "")))));
            }
            else
            {
                string embed = null;
                w.Drawing dr = sdtElement.Descendants<w.Drawing>().FirstOrDefault();
                d.Blip blip = new d.Blip();
                if (dr != null)
                {
                    blip = dr.Descendants<d.Blip>().FirstOrDefault();
                    if (blip != null)
                        embed = blip.Embed;
                    if (embed != null)
                    {
                        p.IdPartPair idpp = wordDocument.MainDocumentPart.Parts.Where(pa => pa.RelationshipId == embed).FirstOrDefault();
                        if (idpp != null)
                        {
                            string Exits = "";
                            foreach (var item in Embeds)
                                if (ByteArrayCompare(item.Image, field.Value as byte[]))
                                    Exits = item.Name;

                            if (!String.IsNullOrEmpty(Exits))
                                blip.Embed = Exits;
                            else
                            {
                                using (MemoryStream ms = new MemoryStream(field.Value as byte[]))
                                {
                                    p.ImagePart imagePart = wordDocument.MainDocumentPart.AddImagePart(p.ImagePartType.Jpeg);
                                    imagePart.FeedData(ms);
                                    blip.Embed = wordDocument.MainDocumentPart.GetIdOfPart(imagePart);
                                    Embeds.Add(new Embed { Name = blip.Embed, Image = field.Value as byte[] });
                                }
                            }
                        }
                    }
                }
                else
                {
                    w.Text text4edit =
                        sdtElement.Descendants<w.Text>()
                            .FirstOrDefault();
                    if (text4edit != null)
                        text4edit.Text = "{#NA}";
                    else
                        sdtElement.Append(
                            new w.SdtContentBlock(new w.Paragraph(
                                new w.Run(
                                    new w.Text("{#NA}")))));
                }
            }
        }

        private void FillTable(TableContent table, w.SdtElement sdtBlock)
        {
            if (table.Rows.Count() < 1)
                return;

            string[] firstfield = new string[table.Rows.Count()];
            int i = 0;
            foreach (var row in table.Rows)
                firstfield[i++] = row.FirstOrDefault().Name.Trim();

            w.Table table4edit = null;
            if (sdtBlock != null)
                table4edit = sdtBlock.Descendants<w.Table>().FirstOrDefault();

            if (table4edit != null)
            {
                var cellcontent = table4edit.Descendants<w.SdtElement>()
                                    .FirstOrDefault(c =>
                                    {
                                        w.SdtProperties properties = c.Elements<w.SdtProperties>().FirstOrDefault();
                                        if (properties != null)
                                        {
                                            w.SdtAlias alias = properties.Elements<w.SdtAlias>().FirstOrDefault();
                                            if (alias != null && alias.Val.HasValue && firstfield.Contains(alias.Val.Value.Trim()))
                                                return true;
                                        }
                                        return false;
                                    });
                w.TableRow row4edit = null;
                if (cellcontent != null)
                {
                    o.OpenXmlElement element4findrow = cellcontent.Parent;
                    while (row4edit == null)
                    {
                        row4edit = (element4findrow.Parent as w.TableRow);
                        element4findrow = element4findrow.Parent;
                    }

                    foreach (var row in table.Rows)
                    {
                        w.TableRow rowtmp = row4edit.CloneNode(true) as w.TableRow;
                        foreach (var field in row)
                        {
                            var content4edit = rowtmp.Descendants<w.SdtElement>()
                                    .FirstOrDefault(c =>
                                    {
                                        w.SdtProperties properties = c.Elements<w.SdtProperties>().FirstOrDefault();
                                        if (properties != null)
                                        {
                                            w.SdtAlias alias = properties.Elements<w.SdtAlias>().FirstOrDefault();
                                            if (alias != null && alias.Val.HasValue && alias.Val.Value.Trim() == field.Name.Trim())
                                                return true;
                                        }
                                        return false;
                                    });
                            if (content4edit != null)
                                FillField(field, content4edit);
                        }
                        table4edit.InsertBefore(new w.TableRow(rowtmp.OuterXml), row4edit);
                    }
                    table4edit.RemoveChild(row4edit);
                }
            }
        }

        private static bool ByteArrayCompare(byte[] a1, byte[] a2)
        {
            if (a1.Length != a2.Length)
                return false;

            for (int i = 0; i < a1.Length; i++)
                if (a1[i] != a2[i])
                    return false;

            return true;
        }

        private string CreateFileOutput(string fileInput, int? index)
        {
            if (!String.IsNullOrEmpty(fileInput.Trim()))
            {
                index = index.HasValue ? index.Value > 0 ? index.Value : 1 : 1;
                string tmpFileName = String.Format("{0}.{1:000}.docx", fileInput.Substring(0, fileInput.Length - 5), index.Value);

                if (File.Exists(tmpFileName))
                {
                    try
                    {
                        File.Delete(tmpFileName);
                        File.Copy(fileInput, tmpFileName);
                        return tmpFileName;
                    }
                    catch
                    {
                        return CreateFileOutput(fileInput, ++index);
                    }
                }
                else
                {
                    try
                    {
                        File.Copy(fileInput, tmpFileName);
                        return tmpFileName;
                    }
                    catch
                    {
                        return "NULL";
                    }
                }
            }
            else
                return "NULL";
        }
        #endregion

        // Save and exit
        public void Dispose()
        {
            if (wordDocument != null)
                wordDocument.Dispose();
        }
    }
}
