using System.IO.Packaging;
using System.Reflection;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLPresentationSample;

using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;

public class PresentationEditor
{
    private int CountSlides(PresentationDocument ppt, bool includeHidden = true)
    {
        PresentationPart? presentationPart = ppt.PresentationPart;
        var slides = presentationPart?.SlideParts;
        if (!includeHidden)
        {
            slides = slides.Where(
                (s) => (s.Slide != null) &&
                       ((s.Slide.Show == null) || (s.Slide.Show.HasValue &&
                                                   s.Slide.Show.Value)));
        }

        return slides.Count();
    }

    public void MoveSlide(PresentationDocument presentationDocument, int from, int to)
    {
        if (presentationDocument == null)
        {
            throw new ArgumentNullException(nameof(presentationDocument));
        }

        // Call the CountSlides method to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

        // Verify that both from and to positions are within range and different from one another.
        if (from < 0 || from >= slidesCount)
        {
            throw new ArgumentOutOfRangeException(nameof(from));
        }

        if (to < 0 || from >= slidesCount || to == from)
        {
            throw new ArgumentOutOfRangeException(nameof(to));
        }

        // Get the presentation part from the presentation document.
        PresentationPart? presentationPart = presentationDocument.PresentationPart;

        // The slide count is not zero, so the presentation must contain slides.            
        Presentation? presentation = presentationPart?.Presentation;
        SlideIdList? slideIdList = presentation?.SlideIdList;

        // Get the slide ID of the source slide.
        SlideId? sourceSlide = slideIdList?.ChildElements[from] as SlideId;

        SlideId? targetSlide = null;

        // Identify the position of the target slide after which to move the source slide.
        if (to == 0)
        {
            targetSlide = null;
        }

        if (from < to)
        {
            targetSlide = slideIdList?.ChildElements[to] as SlideId;
        }
        else
        {
            targetSlide = slideIdList?.ChildElements[to - 1] as SlideId;
        }

        // Remove the source slide from its current position.
        sourceSlide?.Remove();

        // Insert the source slide at its new position after the target slide.
        slideIdList?.InsertAfter(sourceSlide, targetSlide);
    }

    private SlideId? GetPageSlideId(PresentationDocument presentationDocument, int page)
    {
        if (presentationDocument == null)
        {
            throw new ArgumentNullException(nameof(presentationDocument));
        }

        // Call the CountSlides method to get the number of slides in the presentation.
        int slidesCount = CountSlides(presentationDocument);

        // Verify that both from and to positions are within range and different from one another.
        if (page < 0 || page >= slidesCount)
        {
            throw new ArgumentOutOfRangeException(nameof(page));
        }

        // Get the presentation part from the presentation document.
        var presentationPart = presentationDocument.PresentationPart;

        // The slide count is not zero, so the presentation must contain slides.            
        var presentation = presentationPart?.Presentation;
        var slideIdList = presentation?.SlideIdList;

        // Get the slide ID of the source slide.
        return slideIdList?.ChildElements[page] as SlideId;
    }

    private void Copy(string docName, string newDocName)
    {
        // var presentationPackage = Package.Open(docName, FileMode.Open, FileAccess.Read);
        // using var ppt = PresentationDocument.Open(presentationPackage);
        using var ppt = PresentationDocument.Open(docName, false);
        using var newPpt = ppt.SaveAs(newDocName);
        ppt.Close();
        newPpt.Close();
    }

    private SlidePart GetSlidePartFromSlideId(PresentationDocument presentationDocument, SlideId slideId)
    {
        PresentationPart part = presentationDocument.PresentationPart;
        string relId = slideId.RelationshipId;

        // Get the slide part by the relationship ID.
        return (SlidePart) part.GetPartById(relId);
    }

    private void EmbedImage(PresentationDocument presentationDocument, int page, Stream image, Drawing.Offset offset,
        Drawing.Extents extents)
    {
        var slidePart = GetSlidePartFromSlideId(presentationDocument, GetPageSlideId(presentationDocument, page));
        ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg);
        imagePart.FeedData(image);


        var tree = slidePart.Slide.Descendants<ShapeTree>().First();

        var picture = new Picture();
        picture.NonVisualPictureProperties = new NonVisualPictureProperties();
        picture.NonVisualPictureProperties.Append(new NonVisualDrawingProperties
        {
            Name = "Image Shape",
            Id = (UInt32) tree.ChildElements.Count - 1
        });

        var nonVisualPictureDrawingProperties = new NonVisualPictureDrawingProperties();
        nonVisualPictureDrawingProperties.Append(new Drawing.PictureLocks()
        {
            NoChangeAspect = true
        });
        picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        picture.NonVisualPictureProperties.Append(new ApplicationNonVisualDrawingProperties());

        var blipFill = new BlipFill();
        var blip1 = new Drawing.Blip()
        {
            Embed = slidePart.GetIdOfPart(imagePart)
        };
        var blipExtensionList1 = new Drawing.BlipExtensionList();
        var blipExtension1 = new Drawing.BlipExtension()
        {
            Uri = $"{{{Guid.NewGuid().ToString()}}}"
        };
        var useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi()
        {
            Val = false
        };
        useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
        blipExtension1.Append(useLocalDpi1);
        blipExtensionList1.Append(blipExtension1);
        blip1.Append(blipExtensionList1);
        var stretch = new Drawing.Stretch();
        stretch.Append(new Drawing.FillRectangle());
        blipFill.Append(blip1);
        blipFill.Append(stretch);
        picture.Append(blipFill);

        picture.ShapeProperties = new ShapeProperties();
        picture.ShapeProperties.Transform2D = new Drawing.Transform2D();
        picture.ShapeProperties.Transform2D.Append(offset);
        picture.ShapeProperties.Transform2D.Append(extents);
        picture.ShapeProperties.Append(new Drawing.PresetGeometry
        {
            Preset = Drawing.ShapeTypeValues.Rectangle
        });

        tree.Append(picture);
    }

    public void AppendText(PresentationDocument presentationDocument, int page, string text, Drawing.Offset offset,
        Drawing.Extents extents)
    {
        var slidePart = GetSlidePartFromSlideId(presentationDocument, GetPageSlideId(presentationDocument, page));
        var tree = slidePart.Slide.Descendants<ShapeTree>().First();
        var textShape = tree.AppendChild(new Shape());
        textShape.NonVisualShapeProperties = new NonVisualShapeProperties
        (new NonVisualDrawingProperties() {Id = (UInt32) tree.ChildElements.Count - 1, Name = "My Text"},
            new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() {NoGrouping = true}),
            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() {Type = PlaceholderValues.Body}));
        textShape.ShapeProperties = new ShapeProperties();
        // Specify the text of the title shape.
        var paragraphProperties = new Drawing.ParagraphProperties()
        {
            Alignment = Drawing.TextAlignmentTypeValues.Left
            // FontAlignment = Drawing.TextFontAlignmentValues.Top
        };
        paragraphProperties.Append(new Drawing.CharacterBullet()
        {
            Char = "-"
        });
        paragraphProperties.Append(new Drawing.BulletFont()
        {
            Typeface = "Arial"
        });
        textShape.TextBody = new TextBody(new Drawing.BodyProperties()
            {
                Anchor = Drawing.TextAnchoringTypeValues.Top
            },
            new Drawing.ListStyle(),
            new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() {Text = text}))
            {
                ParagraphProperties = paragraphProperties
            });
        textShape.ShapeProperties.Transform2D = new Drawing.Transform2D();
        textShape.ShapeProperties.Append(new Drawing.PresetGeometry
        {
            Preset = Drawing.ShapeTypeValues.Rectangle
        });
        textShape.ShapeProperties.Transform2D.Append(offset);
        textShape.ShapeProperties.Transform2D.Append(extents);
        var outline = new Drawing.Outline();
        outline.Append(new Drawing.SolidFill()
        {
            RgbColorModelHex = new Drawing.RgbColorModelHex()
            {
                Val = new HexBinaryValue("FFFFFF")
            }
        });
        textShape.ShapeProperties.Append(outline);
    }

    private Drawing.TableCell CreateTextCell(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            text = string.Empty;
        }

        // Declare and instantiate the table cell
        // Create table cell with the below order:
        // a:tc(TableCell)->a:txbody(TextBody)->a:p(Paragraph)->a:r(Run)->a:t(Text)
        var tableCell = new Drawing.TableCell();

        //  Declare and instantiate the text body
        var textBody = new Drawing.TextBody();
        var bodyProperties = new Drawing.BodyProperties();
        var listStyle = new Drawing.ListStyle();

        var paragraph = new Drawing.Paragraph();
        var run = new Drawing.Run();
        var runProperties = new Drawing.RunProperties()
            {Language = "en-US", Dirty = false};
        var text2 = new Drawing.Text();
        text2.Text = text;
        run.Append(runProperties);
        run.Append(text2);
        var endParagraphRunProperties =
            new Drawing.EndParagraphRunProperties() {Language = "en-US", Dirty = false};

        paragraph.Append(run);
        paragraph.Append(endParagraphRunProperties);
        textBody.Append(bodyProperties);
        textBody.Append(listStyle);
        textBody.Append(paragraph);

        var tableCellProperties = new Drawing.TableCellProperties();
        tableCell.Append(textBody);
        tableCell.Append(tableCellProperties);

        return tableCell;
    }

    private Drawing.Table GenerateTable()
    {
        // Declare and instantiate table
        var table = new Drawing.Table();

        // Specify the required table properties for the table
        var tableProperties = new Drawing.TableProperties()
            {FirstRow = true, BandRow = true, BandColumn = true};
        var tableStyleId = new Drawing.TableStyleId();
        tableStyleId.Text = $"{{{Guid.NewGuid().ToString()}}}";

        tableProperties.Append(tableStyleId);

        // Declare and instantiate tablegrid and columns depending on your columns
        var tableGrid1 = new Drawing.TableGrid();
        foreach (var column in Enumerable.Range(0, 6))
        {
            var gridColumn = new Drawing.GridColumn() {Width = 1948000L};
            tableGrid1.Append(gridColumn);
        }

        table.Append(tableProperties);
        table.Append(tableGrid1);

        var headerRow = new Drawing.TableRow() {Height = 370840L};
        foreach (var column in Enumerable.Range(0, 5).Select(r => r))
        {
            headerRow.Append(CreateTextCell($"Header Column:{column}"));
        }

        headerRow.Append(CreateTextCell($"Remarks"));
        table.Append(headerRow);

        foreach (var row in Enumerable.Range(0, 5).Select(r => r.ToString()))
        {
            var bodyRow = new Drawing.TableRow() {Height = 370840L};
            foreach (var column in Enumerable.Range(0, 5).Select(r => r.ToString()))
            {
                bodyRow.Append(CreateTextCell($"Body Row:{row} Column:{column}"));
            }

            table.Append(bodyRow);
        }

        return table;
    }

    private void AppendTable(PresentationDocument presentationDocument, int page)
    {
        var tableSlidePart = GetSlidePartFromSlideId(presentationDocument, GetPageSlideId(presentationDocument, page));
        var tree = tableSlidePart.Slide.Descendants<ShapeTree>().First();
        var graphicFrame = tree.AppendChild
        (new GraphicFrame(new NonVisualGraphicFrameProperties
        (new NonVisualDrawingProperties()
            {
                Name = "Table Shape",
                Id = (UInt32) tree.ChildElements.Count - 1
            },
            new NonVisualGraphicFrameDrawingProperties(),
            new ApplicationNonVisualDrawingProperties())));

        var offset = new Drawing.Offset() {X = 250000L, Y = 2000000L};
        graphicFrame.Transform = new Transform(offset);
        graphicFrame.Graphic = new Drawing.Graphic(new Drawing.GraphicData(GenerateTable())
            {Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"});
    }

    private void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle)
    {
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        if (slideTitle == null)
        {
            throw new ArgumentNullException("slideTitle");
        }

        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // Verify that the presentation is not empty.
        if (presentationPart == null)
        {
            throw new InvalidOperationException("The presentation document is empty.");
        }

        // Declare and instantiate a new slide.
        Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
        uint drawingObjectId = 1;

        // Construct the slide content.            
        // Specify the non-visual properties of the new slide.
        NonVisualGroupShapeProperties nonVisualProperties =
            slide.CommonSlideData.ShapeTree.AppendChild(new NonVisualGroupShapeProperties());
        nonVisualProperties.NonVisualDrawingProperties = new NonVisualDrawingProperties() {Id = 1, Name = ""};
        nonVisualProperties.NonVisualGroupShapeDrawingProperties = new NonVisualGroupShapeDrawingProperties();
        nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

        // Specify the group shape properties of the new slide.
        slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

        // Declare and instantiate the title shape of the new slide.
        Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());

        drawingObjectId++;

        // Specify the required shape properties for the title shape. 
        titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
        (new NonVisualDrawingProperties() {Id = drawingObjectId, Name = "Title"},
            new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() {NoGrouping = true}),
            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() {Type = PlaceholderValues.Title}));
        titleShape.ShapeProperties = new ShapeProperties();

        // Specify the text of the title shape.
        titleShape.TextBody = new TextBody(new Drawing.BodyProperties(),
            new Drawing.ListStyle(),
            new Drawing.Paragraph(new Drawing.Run(new Drawing.Text() {Text = slideTitle})));

        // Declare and instantiate the body shape of the new slide.
        Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
        drawingObjectId++;

        // Specify the required shape properties for the body shape.
        bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties() {Id = drawingObjectId, Name = "Content Placeholder"},
            new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() {NoGrouping = true}),
            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() {Index = 1}));
        bodyShape.ShapeProperties = new ShapeProperties();

        // Specify the text of the body shape.
        bodyShape.TextBody = new TextBody(new Drawing.BodyProperties(),
            new Drawing.ListStyle(),
            new Drawing.Paragraph());

        // Create the slide part for the new slide.
        SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

        // Save the new slide part.
        slide.Save(slidePart);

        // Modify the slide ID list in the presentation part.
        // The slide ID list should not be null.
        SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

        // Find the highest slide ID in the current list.
        uint maxSlideId = 1;
        SlideId prevSlideId = null;

        foreach (SlideId slideId in slideIdList.ChildElements)
        {
            if (slideId.Id > maxSlideId)
            {
                maxSlideId = slideId.Id;
            }

            position--;
            if (position == 0)
            {
                prevSlideId = slideId;
            }
        }

        maxSlideId++;

        // Get the ID of the previous slide.
        SlidePart lastSlidePart;

        if (prevSlideId != null)
        {
            lastSlidePart = (SlidePart) presentationPart.GetPartById(prevSlideId.RelationshipId);
        }
        else
        {
            lastSlidePart =
                (SlidePart) presentationPart.GetPartById(((SlideId) (slideIdList.ChildElements[0])).RelationshipId);
        }

        // Use the same slide layout as that of the previous slide.
        if (null != lastSlidePart.SlideLayoutPart)
        {
            slidePart.AddPart(lastSlidePart.SlideLayoutPart);
        }

        // Insert the new slide into the slide list after the previous slide.
        SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
        newSlideId.Id = maxSlideId;
        newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

        // Save the modified presentation.
        presentationPart.Presentation.Save();
    }

    public void DoEdit(string docName, string newDocName, Stream image)
    {
        Copy(docName, newDocName);
        using var ppt = PresentationDocument.Open(newDocName, true);


        var offset = new Drawing.Offset
        {
            X = 7500000,
            Y = 2000000,
        };
        var extents = new Drawing.Extents
        {
            Cx = 3000000,
            Cy = 3000000,
        };
        EmbedImage(ppt, 1, image, offset, extents);

        var textOffset = new Drawing.Offset
        {
            X = 2000000,
            Y = 2000000,
        };
        var textExtents = new Drawing.Extents
        {
            Cx = 3000000,
            Cy = 5000000,
        };
        AppendText(ppt, 2, "Hello\nWorld", textOffset, textExtents);
        MoveSlide(ppt, 2, 1);

        InsertNewSlide(ppt, 3, "New Slide");
        AppendTable(ppt, 3);

        ppt.Save();
        ppt.Close();
        Console.WriteLine($"Saved as {newDocName}");
    }
}