using BatchAndReport.DAO;
using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Services;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

public class DgaEsignService
{
    private readonly IConverter _pdfConverter;
    public DgaEsignService(
      IConverter pdfConverter
    )
    {

        _pdfConverter = pdfConverter;
    }
    public  async Task RegisterDocument(string detail)
    {
       
    }


}
