using System.Drawing;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dox2Word.Logging;
using DocumentFormat.OpenXml.Drawing;
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
            ImagePartType? type = System.IO.Path.GetExtension(path) switch
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

            var element = new Drawing(
                new DW.Inline()
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U,

                    Extent = new DW.Extent()
                    {
                        Cx = widthEmus,
                        Cy = heightEmus
                    },
                    EffectExtent = new DW.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    DocProperties = new DW.DocProperties()
                    {
                        Id = this.id,
                        Name = $"Picture {this.id}"
                    },
                    NonVisualGraphicFrameDrawingProperties = new DW.NonVisualGraphicFrameDrawingProperties()
                    {
                        GraphicFrameLocks = new GraphicFrameLocks()
                        {
                            NoChangeAspect = true
                        },
                    },
                    Graphic = new Graphic()
                    {
                        GraphicData = new GraphicData(
                            new PIC.Picture()
                            {
                                NonVisualPictureProperties = new PIC.NonVisualPictureProperties()
                                {
                                    NonVisualDrawingProperties = new PIC.NonVisualDrawingProperties()
                                    {
                                        Id = this.id,
                                        Name = $"Picture {this.id}"
                                    },
                                    NonVisualPictureDrawingProperties = new PIC.NonVisualPictureDrawingProperties(),
                                },
                                BlipFill = new PIC.BlipFill(
                                    new Blip(
                                        new BlipExtensionList(
                                            new BlipExtension()
                                            {
                                                Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                            })
                                        )
                                    {
                                        Embed = relationshipId,
                                        CompressionState = BlipCompressionValues.Print
                                    },
                                    new Stretch()
                                    {
                                        FillRectangle = new FillRectangle(),
                                    }
                                ),
                                ShapeProperties = new PIC.ShapeProperties(
                                    new Transform2D()
                                    {
                                        Offset = new Offset() { X = 0L, Y = 0L },
                                        Extents = new Extents() { Cx = widthEmus, Cy = heightEmus },
                                    },
                                    new PresetGeometry()
                                    {
                                        AdjustValueList = new AdjustValueList(),
                                        Preset = ShapeTypeValues.Rectangle,
                                    }),
                                }
                            )
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }
                    }
                }
            );

            this.id++;

            return element;
        }
    }
}
