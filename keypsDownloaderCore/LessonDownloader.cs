using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reactive.Subjects;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using keypsDownloaderCore.Models;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.Metadata;
using SixLabors.ImageSharp.PixelFormats;
using D = DocumentFormat.OpenXml.Drawing;


namespace keypsDownloaderCore {
    public class LessonDownloader {
        private string DownloadFolder { get; set; }

        public LessonDownloader(string downloadFolder) {
            DownloadedLessons = new Subject<Lesson>();
        }

        public Subject<Lesson> DownloadedLessons { get; }

        public async Task DownloadLessonsAsync(IEnumerable<Lesson> lessonList) {
            var taskList = lessonList.Select(lesson => DownloadLesson(lesson)).ToList();

            await Task.WhenAll(taskList);
        }

        public async Task DownloadLesson(Lesson lesson) {
            var pages = await GetListOfSlidePages(lesson);
            await CreatePresentationFromPages(lesson, await DownloadListOfPagesAsync(pages));
            DownloadedLessons.OnNext(lesson);
        }

        private async Task<IEnumerable<Page>> GetListOfSlidePages(Lesson lesson) {
            return await Task.Run(() => {
                var driverService = ChromeDriverService.CreateDefaultService();
                driverService.HideCommandPromptWindow = true;
                
                var chromeOptions = new ChromeOptions();
                chromeOptions.AddArguments("headless", "--disable-gpu");


                ChromeDriver driver = new ChromeDriver(driverService, chromeOptions);
                // ChromeDriver driver = new ChromeDriver();
                driver.Navigate().GoToUrl(lesson.Url);
                WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 1, 0));
                var slides = wait.Until(d => {
                    try {
                        var elementToBeDisplayed = driver.FindElementsByClassName("slide");
                        return elementToBeDisplayed.Count > 0 ? elementToBeDisplayed : null;
                    }
                    catch (StaleElementReferenceException) {
                        return null;
                    }
                    catch (NoSuchElementException) {
                        return null;
                    }
                });

                // var slides = _driver.FindElementsByClassName("slide");
                List<Page> list = new List<Page>();
                foreach (IWebElement page in slides) {
                    var slideUrl = $"{lesson.BaseKapittaUrl}/{page.GetAttribute("xlink:href")}";
                    var slideId = page.GetAttribute("id");
                    var pageNumber = 0;
                    if (!page.Equals(slides[0])) {
                        var match = Regex.Match(slideId, @"image(?<page_number>\d+)").Groups["page_number"];
                        // var match = Regex.Match(newUrl, @"slide-(?<page_number>\d+).png").Groups["page_number"];
                        // // Bazen deskshare.png vs oluyor burada, onda da boş değer dönüyor. Onları ekarte etmek için bu satır var.
                        if (match.Value == "") continue;
                        pageNumber = int.Parse(match.Value);
                    }

                    list.Add(new Page(url: slideUrl, number: pageNumber));
                }

                driver.Close();

                // if (list.Find(page => page.PageNumber == 1) == null) {
                //     throw new Exception("İndirilen slayt 2 parça ve bu 2.si");
                // }

                var distinctList = list.GroupBy(p => p.PageNumber).Select(g => g.First());

                Debug.WriteLine("Fetched slide list.");
                return distinctList;
            });
        }

        private async Task<IEnumerable<Page>> DownloadListOfPagesAsync(IEnumerable<Page> pages) {
            Debug.WriteLine("Downloading pages.");
            HttpClient httpClient = new HttpClient();
            var pageList = pages.ToList();
            foreach (Page page in pageList) {
                page.FileName = Path.GetTempFileName();
                await using (FileStream destinationStream =
                    File.Open(page.FileName, FileMode.OpenOrCreate, FileAccess.ReadWrite)) {
                    var sourceStream = await httpClient.GetStreamAsync(page.Url);
                    await sourceStream.CopyToAsync(destinationStream);
                }

                Debug.WriteLine($"Downloaded {page.PageNumber}");
            }

            Debug.WriteLine("Page list downloaded.");
            return pageList;
        }

        async Task CreatePresentationFromPages(Lesson lesson, IEnumerable<Page> pages) {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var presentation = await CreatePresentation($"{desktop}\\{lesson.Name}.pptx");
            var sortedPages = pages.OrderBy(p => p.PageNumber);
            foreach (var item in sortedPages.Select((x, i) => new {Value = x, Index = i})) {
                InsertNewSlide(presentation, item.Index, item.Value);
            }

            presentation.Save();
            presentation.Close();

            Debug.WriteLine("Presentation complete.");
        }

        // Insert the specified slide into the presentation at the specified position.
        public static void InsertNewSlide(PresentationDocument presentationDocument, int position, Page page) {
            if (presentationDocument == null) {
                throw new ArgumentNullException("presentationDocument");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Verify that the presentation is not empty.
            if (presentationPart == null) {
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

            // // Declare and instantiate the title shape of the new slide.
            // Shape titleShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
            //
            // drawingObjectId++;
            //
            // // Specify the required shape properties for the title shape. 
            // titleShape.NonVisualShapeProperties = new NonVisualShapeProperties
            // (new NonVisualDrawingProperties() {Id = drawingObjectId, Name = "Title"},
            //     new NonVisualShapeDrawingProperties(new D.ShapeLocks() {NoGrouping = true}),
            //     new ApplicationNonVisualDrawingProperties(new PlaceholderShape() {Type = PlaceholderValues.Title}));
            // titleShape.ShapeProperties = new ShapeProperties();
            //
            // // Specify the text of the title shape.
            // titleShape.TextBody = new TextBody(new D.BodyProperties(),
            //     new D.ListStyle(),
            //     new D.Paragraph(new D.Run(new D.Text() {Text = ""})));
            //
            // // Declare and instantiate the body shape of the new slide.
            // Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new Shape());
            // drawingObjectId++;
            //
            // // Specify the required shape properties for the body shape.
            // bodyShape.NonVisualShapeProperties = new NonVisualShapeProperties(
            //     new NonVisualDrawingProperties() {Id = drawingObjectId, Name = "Content Placeholder"},
            //     new NonVisualShapeDrawingProperties(new D.ShapeLocks() {NoGrouping = true}),
            //     new ApplicationNonVisualDrawingProperties(new PlaceholderShape() {Index = 1}));
            // bodyShape.ShapeProperties = new ShapeProperties();
            //
            // // Specify the text of the body shape.
            // bodyShape.TextBody = new TextBody(new D.BodyProperties(),
            //     new D.ListStyle(),
            //     new D.Paragraph());

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

            foreach (var openXmlElement in slideIdList.ChildElements) {
                var slideId = (SlideId) openXmlElement;
                if (slideId.Id > maxSlideId) {
                    maxSlideId = slideId.Id;
                }

                position--;
                if (position == 0) {
                    prevSlideId = slideId;
                }
            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null) {
                lastSlidePart = (SlidePart) presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else {
                lastSlidePart =
                    (SlidePart) presentationPart.GetPartById(((SlideId) (slideIdList.ChildElements[0])).RelationshipId);
            }

            // Use the same slide layout as that of the previous slide.
            if (null != lastSlidePart.SlideLayoutPart) {
                slidePart.AddPart(lastSlidePart.SlideLayoutPart);
            }

            AddImageToSlide(slidePart, page);

            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();
        }

        public static void AddImageToSlide(SlidePart slidePart, Page page) {
            var part = slidePart
                .AddImagePart(ImagePartType.Png);

            using (var stream = File.OpenRead(page.FileName)) {
                part.FeedData(stream);
            }

            var tree = slidePart
                .Slide
                .Descendants<ShapeTree>()
                .First();

            var picture = new Picture();

            picture.NonVisualPictureProperties =
                new NonVisualPictureProperties();
            picture.NonVisualPictureProperties.Append(
                new NonVisualDrawingProperties {
                    Name = "My Shape",
                    Id = (UInt32) tree.ChildElements.Count - 1
                });

            var nonVisualPictureDrawingProperties =
                new DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties();
            nonVisualPictureDrawingProperties.Append(new D.PictureLocks() {
                NoChangeAspect = true
            });
            picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
            picture.NonVisualPictureProperties.Append(
                new ApplicationNonVisualDrawingProperties());

            var blipFill = new DocumentFormat.OpenXml.Presentation.BlipFill();
            var blip1 = new DocumentFormat.OpenXml.Drawing.Blip() {
                Embed = slidePart.GetIdOfPart(part)
            };
            var blipExtensionList1 = new DocumentFormat.OpenXml.Drawing.BlipExtensionList();
            var blipExtension1 = new DocumentFormat.OpenXml.Drawing.BlipExtension() {
                Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
            };
            var useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi() {
                Val = false
            };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            blipExtension1.Append(useLocalDpi1);
            blipExtensionList1.Append(blipExtension1);
            blip1.Append(blipExtensionList1);
            var stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
            stretch.Append(new DocumentFormat.OpenXml.Drawing.FillRectangle());
            blipFill.Append(blip1);
            blipFill.Append(stretch);
            picture.Append(blipFill);

            picture.ShapeProperties = new ShapeProperties();
            picture.ShapeProperties.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();
            picture.ShapeProperties.Transform2D.Append(new D.Offset {
                X = 0,
                Y = 0,
            });
            picture.ShapeProperties.Transform2D.Append(new D.Extents {
                Cx = 9144000, Cy = 6858000
            });
            picture.ShapeProperties.Append(new D.PresetGeometry {
                Preset = D.ShapeTypeValues.Rectangle
            });

            tree.Append(picture);
        }

        public async Task<PresentationDocument> CreatePresentation(string filename) {
            // Create a presentation at a specified file path. The presentation document type is pptx, by default.
            PresentationDocument presentationDoc =
                PresentationDocument.Create(filename, PresentationDocumentType.Presentation);
            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            CreatePresentationParts(presentationPart);

            return presentationDoc;
        }

        private static void CreatePresentationParts(PresentationPart presentationPart) {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId()
                {Id = (UInt32Value) 2147483648U, RelationshipId = "rId1"});
            SlideIdList slideIdList1 =
                new SlideIdList(new SlideId() {Id = (UInt32Value) 256U, RelationshipId = "rId2"});
            SlideSize slideSize1 = new SlideSize() {Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3};
            NotesSize notesSize1 = new NotesSize() {Cx = 6858000, Cy = 9144000};
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1,
                defaultTextStyle1);

            SlidePart slidePart1;
            SlideLayoutPart slideLayoutPart1;
            SlideMasterPart slideMasterPart1;
            ThemePart themePart1;


            slidePart1 = CreateSlidePart(presentationPart);
            slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
            slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
            themePart1 = CreateTheme(slideMasterPart1);

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
            presentationPart.AddPart(slideMasterPart1, "rId1");
            presentationPart.AddPart(themePart1, "rId5");
        }

        private static SlidePart CreateSlidePart(PresentationPart presentationPart) {
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
            slidePart1.Slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""},
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new D.TransformGroup()),
                        new Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties() {Id = (UInt32Value) 2U, Name = "Title 1"},
                                new NonVisualShapeDrawingProperties(new D.ShapeLocks() {NoGrouping = true}),
                                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                            new ShapeProperties(),
                            new TextBody(
                                new D.BodyProperties(),
                                new D.ListStyle(),
                                new D.Paragraph(new D.EndParagraphRunProperties() {Language = "en-US"}))))),
                new ColorMapOverride(new D.MasterColorMapping()));
            return slidePart1;
        }

        private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1) {
            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            SlideLayout slideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""},
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new D.TransformGroup()),
                    new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties() {Id = (UInt32Value) 2U, Name = ""},
                            new NonVisualShapeDrawingProperties(new D.ShapeLocks() {NoGrouping = true}),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                        new ShapeProperties(),
                        new TextBody(
                            new D.BodyProperties(),
                            new D.ListStyle(),
                            new D.Paragraph(new D.EndParagraphRunProperties()))))),
                new ColorMapOverride(new D.MasterColorMapping()));
            slideLayoutPart1.SlideLayout = slideLayout;
            return slideLayoutPart1;
        }

        private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1) {
            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
            SlideMaster slideMaster = new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties() {Id = (UInt32Value) 1U, Name = ""},
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new D.TransformGroup()),
                    new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties() {Id = (UInt32Value) 2U, Name = "Title Placeholder 1"},
                            new NonVisualShapeDrawingProperties(new D.ShapeLocks() {NoGrouping = true}),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape()
                                {Type = PlaceholderValues.Title})),
                        new ShapeProperties(),
                        new TextBody(
                            new D.BodyProperties(),
                            new D.ListStyle(),
                            new D.Paragraph())))),
                new ColorMap() {
                    Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1,
                    Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2,
                    Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2,
                    Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4,
                    Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6,
                    Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                },
                new SlideLayoutIdList(new SlideLayoutId() {Id = (UInt32Value) 2147483649U, RelationshipId = "rId1"}),
                new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
            slideMasterPart1.SlideMaster = slideMaster;

            return slideMasterPart1;
        }

        private static ThemePart CreateTheme(SlideMasterPart slideMasterPart1) {
            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
            D.Theme theme1 = new D.Theme() {Name = "Office Theme"};

            D.ThemeElements themeElements1 = new D.ThemeElements(
                new D.ColorScheme(
                    new D.Dark1Color(new D.SystemColor() {Val = D.SystemColorValues.WindowText, LastColor = "000000"}),
                    new D.Light1Color(new D.SystemColor() {Val = D.SystemColorValues.Window, LastColor = "FFFFFF"}),
                    new D.Dark2Color(new D.RgbColorModelHex() {Val = "1F497D"}),
                    new D.Light2Color(new D.RgbColorModelHex() {Val = "EEECE1"}),
                    new D.Accent1Color(new D.RgbColorModelHex() {Val = "4F81BD"}),
                    new D.Accent2Color(new D.RgbColorModelHex() {Val = "C0504D"}),
                    new D.Accent3Color(new D.RgbColorModelHex() {Val = "9BBB59"}),
                    new D.Accent4Color(new D.RgbColorModelHex() {Val = "8064A2"}),
                    new D.Accent5Color(new D.RgbColorModelHex() {Val = "4BACC6"}),
                    new D.Accent6Color(new D.RgbColorModelHex() {Val = "F79646"}),
                    new D.Hyperlink(new D.RgbColorModelHex() {Val = "0000FF"}),
                    new D.FollowedHyperlinkColor(new D.RgbColorModelHex() {Val = "800080"})) {Name = "Office"},
                new D.FontScheme(
                    new D.MajorFont(
                        new D.LatinFont() {Typeface = "Calibri"},
                        new D.EastAsianFont() {Typeface = ""},
                        new D.ComplexScriptFont() {Typeface = ""}),
                    new D.MinorFont(
                        new D.LatinFont() {Typeface = "Calibri"},
                        new D.EastAsianFont() {Typeface = ""},
                        new D.ComplexScriptFont() {Typeface = ""})) {Name = "Office"},
                new D.FormatScheme(
                    new D.FillStyleList(
                        new D.SolidFill(new D.SchemeColor() {Val = D.SchemeColorValues.PhColor}),
                        new D.GradientFill(
                            new D.GradientStopList(
                                new D.GradientStop(new D.SchemeColor(new D.Tint() {Val = 50000},
                                            new D.SaturationModulation() {Val = 300000})
                                        {Val = D.SchemeColorValues.PhColor})
                                    {Position = 0},
                                new D.GradientStop(new D.SchemeColor(new D.Tint() {Val = 37000},
                                            new D.SaturationModulation() {Val = 300000})
                                        {Val = D.SchemeColorValues.PhColor})
                                    {Position = 35000},
                                new D.GradientStop(new D.SchemeColor(new D.Tint() {Val = 15000},
                                            new D.SaturationModulation() {Val = 350000})
                                        {Val = D.SchemeColorValues.PhColor})
                                    {Position = 100000}
                            ),
                            new D.LinearGradientFill() {Angle = 16200000, Scaled = true}),
                        new D.NoFill(),
                        new D.PatternFill(),
                        new D.GroupFill()),
                    new D.LineStyleList(
                        new D.Outline(
                            new D.SolidFill(
                                new D.SchemeColor(
                                    new D.Shade() {Val = 95000},
                                    new D.SaturationModulation() {Val = 105000}) {Val = D.SchemeColorValues.PhColor}),
                            new D.PresetDash() {Val = D.PresetLineDashValues.Solid}) {
                            Width = 9525,
                            CapType = D.LineCapValues.Flat,
                            CompoundLineType = D.CompoundLineValues.Single,
                            Alignment = D.PenAlignmentValues.Center
                        },
                        new D.Outline(
                            new D.SolidFill(
                                new D.SchemeColor(
                                    new D.Shade() {Val = 95000},
                                    new D.SaturationModulation() {Val = 105000}) {Val = D.SchemeColorValues.PhColor}),
                            new D.PresetDash() {Val = D.PresetLineDashValues.Solid}) {
                            Width = 9525,
                            CapType = D.LineCapValues.Flat,
                            CompoundLineType = D.CompoundLineValues.Single,
                            Alignment = D.PenAlignmentValues.Center
                        },
                        new D.Outline(
                            new D.SolidFill(
                                new D.SchemeColor(
                                    new D.Shade() {Val = 95000},
                                    new D.SaturationModulation() {Val = 105000}) {Val = D.SchemeColorValues.PhColor}),
                            new D.PresetDash() {Val = D.PresetLineDashValues.Solid}) {
                            Width = 9525,
                            CapType = D.LineCapValues.Flat,
                            CompoundLineType = D.CompoundLineValues.Single,
                            Alignment = D.PenAlignmentValues.Center
                        }),
                    new D.EffectStyleList(
                        new D.EffectStyle(
                            new D.EffectList(
                                new D.OuterShadow(
                                    new D.RgbColorModelHex(
                                        new D.Alpha() {Val = 38000}) {Val = "000000"}) {
                                    BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false
                                })),
                        new D.EffectStyle(
                            new D.EffectList(
                                new D.OuterShadow(
                                    new D.RgbColorModelHex(
                                        new D.Alpha() {Val = 38000}) {Val = "000000"}) {
                                    BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false
                                })),
                        new D.EffectStyle(
                            new D.EffectList(
                                new D.OuterShadow(
                                    new D.RgbColorModelHex(
                                        new D.Alpha() {Val = 38000}) {Val = "000000"}) {
                                    BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false
                                }))),
                    new D.BackgroundFillStyleList(
                        new D.SolidFill(new D.SchemeColor() {Val = D.SchemeColorValues.PhColor}),
                        new D.GradientFill(
                            new D.GradientStopList(
                                new D.GradientStop(
                                    new D.SchemeColor(new D.Tint() {Val = 50000},
                                            new D.SaturationModulation() {Val = 300000})
                                        {Val = D.SchemeColorValues.PhColor}) {Position = 0},
                                new D.GradientStop(
                                    new D.SchemeColor(new D.Tint() {Val = 50000},
                                            new D.SaturationModulation() {Val = 300000})
                                        {Val = D.SchemeColorValues.PhColor}) {Position = 0},
                                new D.GradientStop(
                                    new D.SchemeColor(new D.Tint() {Val = 50000},
                                            new D.SaturationModulation() {Val = 300000})
                                        {Val = D.SchemeColorValues.PhColor}) {Position = 0}),
                            new D.LinearGradientFill() {Angle = 16200000, Scaled = true}),
                        new D.GradientFill(
                            new D.GradientStopList(
                                new D.GradientStop(
                                    new D.SchemeColor(new D.Tint() {Val = 50000},
                                            new D.SaturationModulation() {Val = 300000})
                                        {Val = D.SchemeColorValues.PhColor}) {Position = 0},
                                new D.GradientStop(
                                    new D.SchemeColor(new D.Tint() {Val = 50000},
                                            new D.SaturationModulation() {Val = 300000})
                                        {Val = D.SchemeColorValues.PhColor}) {Position = 0}),
                            new D.LinearGradientFill() {Angle = 16200000, Scaled = true}))) {Name = "Office"});

            theme1.Append(themeElements1);
            theme1.Append(new D.ObjectDefaults());
            theme1.Append(new D.ExtraColorSchemeList());

            themePart1.Theme = theme1;
            return themePart1;
        }
    }
}