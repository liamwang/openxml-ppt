using Anet.OpenXml.PPT.Defines;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Anet.OpenXml.PPT
{
    public static class PresentationDocumentExtensions
    {
        /// <summary>
        /// 通过页码获取指页
        /// </summary>
        /// <param name="pageNumber">页码（从1开始）</param>
        /// <returns></returns>
        public static SlidePart GetSlideByPageNum(this PresentationDocument document, int pageNumber)
        {
            // 注意：Presentation.SlideIdList 是顺序排列的，
            // 不能使用 ppt.PresentationPart.SlideParts，它是无序的。

            var idList = document.PresentationPart.Presentation.SlideIdList;

            if (idList.Count() < pageNumber)
                return null;

            var slideId = idList.ElementAt(pageNumber - 1) as SlideId;

            var slidePart = document.PresentationPart.GetPartById(slideId.RelationshipId);

            return slidePart as SlidePart;
        }

        public static IEnumerable<SlidePart> GetSlidePartsInOrder(this PresentationPart presentationPart)
        {
            var slideIdList = presentationPart.Presentation.SlideIdList;

            return slideIdList.ChildElements
                .Cast<SlideId>()
                .Select(x => presentationPart.GetPartById(x.RelationshipId))
                .Cast<SlidePart>();
        }

        public static SlidePart CloneSlide(this PresentationDocument document, SlidePart templatePart, int prePageNum)
        {
            // find the presentationPart: makes the API more fluent
            var presentationPart = templatePart.GetParentParts()
                .OfType<PresentationPart>()
                .Single();

            // clone slide contents
            Slide currentSlide = (Slide)templatePart.Slide.CloneNode(true);
            var slidePartClone = presentationPart.AddNewPart<SlidePart>();
            currentSlide.Save(slidePartClone);

            // copy layout part
            slidePartClone.AddPart(templatePart.SlideLayoutPart);

            // append slide

            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // find the highest id
            uint maxSlideId = slideIdList.ChildElements
                .Cast<SlideId>()
                .Max(x => x.Id.Value);

            SlideId preSlideId = slideIdList.ChildElements
                .Cast<SlideId>()
                .ElementAt(prePageNum - 1);

            // Insert the new slide into the slide list after the previous slide.
            var id = maxSlideId + 1;

            SlideId newSlideId = new SlideId();
            slideIdList.InsertAfter(newSlideId, preSlideId);
            newSlideId.Id = id;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePartClone);

            return slidePartClone;
        }

        public static SlidePart CreateSlidePart(this PresentationDocument document, string id)
        {
            var slidePart = document.PresentationPart.AddNewPart<SlidePart>(id);

            slidePart.Slide = new Slide(CommonSlideDefine.NewBlankCommonSlideData());
            slidePart.Slide.CommonSlideData.ShapeTree.Append(
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = 2U, Name = "Title" },
                        new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                        new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                    new P.ShapeProperties(),
                    new P.TextBody(
                        new BodyProperties(),
                        new ListStyle(),
                        new Paragraph()
                    )
                )
            );
            slidePart.Slide.Append(new ColorMapOverride(new MasterColorMapping()));

            return slidePart;
        }

        public static void InsertSlidePart(this PresentationDocument document, SlidePart slidePart, int position)
        {
            var newSlidePart = document.PresentationPart.AddNewPart<SlidePart>("sld59");
            newSlidePart.FeedData(slidePart.GetStream(FileMode.Open));
            //make sure the new slide references the proper slide layout
            newSlidePart.AddPart(slidePart.SlideLayoutPart);
            //SlideIdList slideIdList = document.PresentationPart.Presentation.SlideIdList;
            //uint maxSlideId = 1;
            //SlideId prevSlideId = null;

            //foreach (SlideId slideId in slideIdList.ChildElements)
            //{
            //    if (slideId.Id > maxSlideId)
            //    {
            //        maxSlideId = slideId.Id;
            //    }

            //    position--;
            //    if (position == 0)
            //    {
            //        prevSlideId = slideId;
            //    }

            //}

            //maxSlideId++;


            //SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            //newSlideId.Id = maxSlideId;
            //newSlideId.RelationshipId = document.PresentationPart.GetIdOfPart(slidePart);

            //document.PresentationPart.Presentation.Save();
        }

        public static void DeleteSlide(this PresentationDocument presentationDocument, SlideId slideId)
        {
            // Get the presentation part from the presentation document. 
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Get the presentation from the presentation part.
            Presentation presentation = presentationPart.Presentation;

            // Get the list of slide IDs in the presentation.
            SlideIdList slideIdList = presentation.SlideIdList;

            // Get the relationship ID of the slide.
            string slideRelId = slideId.RelationshipId;

            // Remove the slide from the slide list.
            slideIdList.RemoveChild(slideId);

            // Remove references to the slide from all custom shows.
            if (presentation.CustomShowList != null)
            {
                // Iterate through the list of custom shows.
                foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
                {
                    if (customShow.SlideList != null)
                    {
                        // Declare a link list of slide list entries.
                        LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                        foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                        {
                            // Find the slide reference to remove from the custom show.
                            if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                            {
                                slideListEntries.AddLast(slideListEntry);
                            }
                        }

                        // Remove all references to the slide from the custom show.
                        foreach (SlideListEntry slideListEntry in slideListEntries)
                        {
                            customShow.SlideList.RemoveChild(slideListEntry);
                        }
                    }
                }
            }

            // Save the modified presentation.
            //presentation.Save();

            // Get the slide part for the specified slide.
            SlidePart slidePart = presentationPart.GetPartById(slideRelId) as SlidePart;

            // Remove the slide part.
            presentationPart.DeletePart(slidePart);
        }

        public static void DeleteSlide(this PresentationDocument presentationDocument, int pageNum)
        {
            // Get the presentation part from the presentation document. 
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Get the presentation from the presentation part.
            Presentation presentation = presentationPart.Presentation;

            // Get the slide ID of the specified slide
            SlideId slideId = presentation.SlideIdList.ChildElements[pageNum - 1] as SlideId;

            presentationDocument.DeleteSlide(slideId);
        }
    }
}
