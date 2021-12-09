using System.Drawing;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dox2Word.Logging;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Dox2Word.Generator
{
    public class ImageManager
    {
        private const double MaxWidthCm = 17;
        private static readonly Logger logger = Logger.Instance;

        private readonly MainDocumentPart mainPart;

        private uint id = 1;

        public ImageManager(MainDocumentPart mainPart)
        {
            this.mainPart = mainPart;
        }

        public OpenXmlElement CreateImage(byte[] data, ImagePartType imageType)
        {
            var imagePart = this.mainPart.AddImagePart(imageType);
            imagePart.FeedData(new MemoryStream(data));

            var (widthEmus, heightEmus) = this.CalculateExtents(data);

            return this.CreateImageElement(this.mainPart.GetIdOfPart(imagePart), widthEmus, heightEmus);
        }

        public OpenXmlElement? CreateImage(string path)
        {
            ImagePartType? type = Path.GetExtension(path) switch
            {
                ".png" => ImagePartType.Png,
                ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                ".gif" => ImagePartType.Gif,
                ".tiff" => ImagePartType.Tiff,
                ".ico" => ImagePartType.Icon,
                ".pcx" => ImagePartType.Pcx,
                ".emf" => ImagePartType.Emf,
                ".wmf" => ImagePartType.Wmf,
                _ => null,
            };
            if (type == null)
            {
                logger.Warning($"Unknown image extension for '{path}'. Ignoring");
                return null;
            }

            return this.CreateImage(File.ReadAllBytes(path), type.Value);
        }

        private (long width, long height) CalculateExtents(byte[] data)
        {
            // https://stackoverflow.com/a/8083390/1086121

            using var img = new Bitmap(new MemoryStream(data));
            int widthPx = img.Width;
            int heightPx = img.Height;
            float horzRezDpi = img.HorizontalResolution;
            float vertRezDpi = img.VerticalResolution;
            const int emusPerInch = 914400;
            const int emusPerCm = 360000;
            long widthEmus = (long)((double)widthPx / horzRezDpi * emusPerInch);
            long heightEmus = (long)((double)heightPx / vertRezDpi * emusPerInch);
            long maxWidthEmus = (long)(MaxWidthCm * emusPerCm);
            if (widthEmus > maxWidthEmus)
            {
                decimal ratio = (heightEmus * 1.0m) / widthEmus;
                widthEmus = maxWidthEmus;
                heightEmus = (long)(widthEmus * ratio);
            }

            return (widthEmus, heightEmus);
        }

        private OpenXmlElement CreateImageElement(string relationshipId, long widthEmus, long heightEmus)
        {
            // From https://docs.microsoft.com/en-us/office/open-xml/how-to-insert-a-picture-into-a-word-processing-document
            
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = widthEmus, Cy = heightEmus },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = this.id,
                             Name = $"Picture {this.id}"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = this.id,
                                             Name = $"Picture {this.id}"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                             )
                                         {
                                             Embed = relationshipId,
                                             CompressionState = A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = widthEmus, Cy = heightEmus }),
                                         new A.PresetGeometry()
                                         {
                                             AdjustValueList = new A.AdjustValueList(),
                                             Preset = A.ShapeTypeValues.Rectangle,
                                         })
                                     )
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = 0U,
                         DistanceFromBottom = 0U,
                         DistanceFromLeft = 0U,
                         DistanceFromRight = 0U,
                         EditId = "50D07946"
                     });

            this.id++;

            return element;
        }
    }
}
