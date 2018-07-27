using DocConvert.Helpers;

namespace DocConvert.Models
{
    public enum MediaTypeEnum
    {
        [ConvertDocument]
        Empty,

        [ConvertDocument(MediaType = "text/html", Extension = "html", ConvertType = "4")]
        Html,

        [ConvertDocument(MediaType = "application/pdf", Extension = "pdf", ConvertType = "12")]
        Pdf,

        [ConvertDocument(MediaType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document", Extension = "docx", ConvertType = "13")]
        Docx,

        [ConvertDocument(MediaType = "multipart/form-data")]
        MultiPart
    }
}