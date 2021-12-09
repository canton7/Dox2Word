using System.Drawing;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dox2Word.Logging;
using DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Dox2Word.Model;

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

        public OpenXmlElement CreateImage(byte[] data, ImagePartType imageType, ImageDimensions dimensions)
        {
            var imagePart = this.mainPart.AddImagePart(imageType);
            imagePart.FeedData(new MemoryStream(data));

            var (widthEmus, heightEmus) = this.CalculateExtents(data, dimensions);

            return this.CreateImageElement(this.mainPart.GetIdOfPart(imagePart), widthEmus, heightEmus);
        }

        public OpenXmlElement? CreateImage(string path, ImageDimensions dimensions)
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

            return this.CreateImage(File.ReadAllBytes(path), type.Value, dimensions);
        }

        private (long width, long height) CalculateExtents(byte[] data, ImageDimensions dimensions)
        {
            // https://stackoverflow.com/a/8083390/1086121

            using var img = new Bitmap(new MemoryStream(data));

            long originalWidthEmus = DimensionToEmus(img.Width, img.HorizontalResolution, default);
            long originalHeightEmus = DimensionToEmus(img.Height, img.VerticalResolution, default);

            long widthEmus;
            long heightEmus;
            if (dimensions.Width == null && dimensions.Height == null)
            {
                (widthEmus, heightEmus) = (originalWidthEmus, originalHeightEmus);
            }
            else if (dimensions.Width != null && dimensions.Height == null)
            {
                widthEmus = DimensionToEmus(img.Width, img.HorizontalResolution, dimensions.Width);
                heightEmus = (long)(originalHeightEmus * ((double)widthEmus / originalWidthEmus));
            }
            else if (dimensions.Width == null && dimensions.Height != null)
            {
                heightEmus = DimensionToEmus(img.Height, img.VerticalResolution, dimensions.Height);
                widthEmus = (long)(originalWidthEmus * ((double)heightEmus / originalHeightEmus));
            }
            else
            {
                widthEmus = DimensionToEmus(img.Width, img.HorizontalResolution, dimensions.Width);
                heightEmus = DimensionToEmus(img.Height, img.VerticalResolution, dimensions.Height);
            }

            const int emusPerInch = 914400;
            const int emusPerCm = 360000;
            long maxWidthEmus = (long)(MaxWidthCm * emusPerCm);
            if (widthEmus > maxWidthEmus)
            {
                double ratio = (double)heightEmus / widthEmus;
                widthEmus = maxWidthEmus;
                heightEmus = (long)(widthEmus * ratio);
            }

            long DimensionToEmus(int sizePx, float resolutionDpi, ImageDimension? dimension)
            {
                return dimension is { } d
                    ? d.Unit switch
                    {
                        ImageDimensionUnit.Px => PxToEmus((int)d.Value, resolutionDpi),
                        ImageDimensionUnit.Cm => CmToEmus(d.Value),
                        ImageDimensionUnit.Inch => InchToEmus(d.Value),
                        ImageDimensionUnit.Percent => (long)(PxToEmus(sizePx, resolutionDpi) * (double)d.Value / 100.0),
                    }
                    : PxToEmus(sizePx, resolutionDpi);
            }

            static long PxToEmus(int sizePx, float resolutionDpi) =>
                (long)((double)sizePx / resolutionDpi * emusPerInch);
            static long CmToEmus(double sizeCm) => (long)(sizeCm * emusPerCm);
            static long InchToEmus(double sizeInch) => (long)(sizeInch * emusPerInch);

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
