namespace BatchAndReport.Models
{
    public class DGAModels
    {

    }

    public class DgaTokenModels
    {
    }
    public class DgaRegisterDocModels
    {
        public byte[]? Content { get; set; }
        public byte[]? Attachment { get; set; }
        public string? Clause { get; set; }
        public string? Link { get; set; }
        public string? Page { get; set; }
        public string? Left { get; set; }
        public string? Bottom { get; set; }

    }

    public class DgaEsignDocumentModels
    {
        public int ID { get; set; }
        public string? WFTypeCode { get; set; }
        public int ContractID { get; set; }
        public string? DGA_TemplateID { get; set; }
        
        public string? DGA_DocumentID { get; set; }
        public string? DGA_SignatureID { get; set; }
        public byte[]? DGA_DocumentDataFile { get; set; }
        public string? DGA_DocumentPathFile { get; set; }
        public string? SignBy { get; set; }
        public string? CreateBy { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime? UpdateDate { get; set; }
    }

    public class DgaEsingConfigModels
    {
        public int ID { get; set; }
        public string ConsumerKey { get; set; } = string.Empty;
        public string ConsumerSecret { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public DateTime CreateDate { get; set; }
        public DateTime? UpdateDate { get; set; }
    }
    public class DgaEsignTemplateModels
    {
        public int ID { get; set; }
        public string ContractType { get; set; } = string.Empty;
        public string DocumentName { get; set; } = string.Empty;
        public string TemplateID { get; set; } = string.Empty;
        public string ConsumerKey { get; set; } = string.Empty;
        public string FlagActive { get; set; } = string.Empty;
        public DateTime CreateDate { get; set; }
    }

    public class DgaEsignUrlModels
    {
        public int ID { get; set; }
        public string ServiceCode { get; set; } = string.Empty;
        public string ServiceName { get; set; } = string.Empty;
        public string Method { get; set; } = string.Empty;
        public string UrlProd { get; set; } = string.Empty;
        public string UrlDev { get; set; } = string.Empty;
        public string Example { get; set; } = string.Empty;
        public DateTime CreateDate { get; set; }
    }

    public class DGADocumentModels
    {
      
        public string DocumentID { get; set; } = string.Empty;
    
    }
    public class DgaEsignModels
    {
        public int ID { get; set; }
        public string WFTypeCode { get; set; } = string.Empty;
        public int ContractID { get; set; }
        public string DGA_DocumentID { get; set; } = string.Empty;
        public string DGA_SignatureID { get; set; } = string.Empty;
        public byte[] DGA_DocumentDataFile { get; set; } = Array.Empty<byte>();
        public string DGA_DocumentPathFile { get; set; } = string.Empty;
        public string SignBy { get; set; } = string.Empty;
        public string CreateBy { get; set; } = string.Empty;
        public DateTime CreateDate { get; set; }
        public DateTime? UpdateDate { get; set; }
    }
}
