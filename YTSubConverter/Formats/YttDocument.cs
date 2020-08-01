﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using Arc.YTSubConverter.Util;

namespace Arc.YTSubConverter.Formats
{
    internal class YttDocument : SubtitleDocument
    {
        private const string ZeroWidthSpace = "\x200B";

        private static readonly Size ReferenceVideoDimensions = new Size(1280, 720);

        public YttDocument(string filePath)
        {
            XmlDocument doc = new XmlDocument { PreserveWhitespace = true };
            doc.Load(filePath);

            VideoDimensions = ReferenceVideoDimensions;

            XmlElement head = (XmlElement)doc.SelectSingleNode("/timedtext/head");
            ReadHead(
                head,
                out List<Line> positions,
                out List<Line> windowStyles,
                out List<Section> pens
            );

            bool awaitingAndroidHack = false;
            foreach (XmlElement elem in doc.SelectNodes("/timedtext/body/p"))
            {
                Line line = ReadLine(elem, positions, windowStyles, pens);
                if (line.Sections.Any(s => ColorUtil.IsDark(s.ForeColor)))
                {
                    awaitingAndroidHack = true;
                }
                else if (awaitingAndroidHack)
                {
                    awaitingAndroidHack = false;

                    Line prevLine = Lines.Last();
                    if (line.Sections.All(s => !ColorUtil.IsDark(s.ForeColor) && s.ForeColor.A == 0 && s.BackColor.A == 0 && s.ShadowColors.Count == 0) &&
                        line.Start == prevLine.Start && line.End == prevLine.End && line.Text == prevLine.Text)
                    {
                        continue;
                    }
                    prevLine.AndroidDarkTextHackAllowed = false;
                }

                Lines.Add(line);
            }

            if (awaitingAndroidHack)
                Lines.Last().AndroidDarkTextHackAllowed = false;

            MergeIdenticallyFormattedSections();
        }

        public YttDocument(SubtitleDocument doc)
            : base(doc)
        {
            Lines.RemoveAll(l => !l.Sections.Any(s => s.Text.Length > 0));
        }

        public override void Save(string filePath)
        {
            CloseGaps();
            MergeSimultaneousLines();
            MergeIdenticallyFormattedSections();
            ApplyEnhancements();
            MergeIdenticallyFormattedSections();
            PreventOverlapFlickering();

            Dictionary<Line, int> positions = ExtractAttributes(Lines, new LinePositionComparer(this));
            Dictionary<Line, int> windowStyles = ExtractAttributes(Lines, new LineAlignmentComparer());
            Dictionary<Section, int> pens = ExtractAttributes(Lines.SelectMany(l => l.Sections), new NormalizedSectionFormatComparer());

            // Use LF instead of CRLF as the latter seems to cause the iOS app to bug out
            using XmlWriter writer = XmlWriter.Create(filePath, new XmlWriterSettings { NewLineChars = "\n" });
            writer.WriteStartElement("timedtext");
            writer.WriteAttributeString("format", "3");

            WriteHead(writer, positions, windowStyles, pens);
            WriteBody(writer, positions, windowStyles, pens);

            writer.WriteEndElement();
        }

        private void ReadHead(XmlElement headElement, out List<Line> positions, out List<Line> windowStyles, out List<Section> pens)
        {
            positions = new List<Line>();
            windowStyles = new List<Line>();
            pens = new List<Section>();

            foreach (XmlElement elem in headElement.ChildNodes.OfType<XmlElement>())
            {
                switch (elem.LocalName)
                {
                    case "wp":
                        AddItemAtIndex(positions, ReadWindowPosition(elem));
                        break;

                    case "ws":
                        AddItemAtIndex(windowStyles, ReadWindowStyle(elem));
                        break;

                    case "pen":
                        AddItemAtIndex(pens, ReadPen(elem));
                        break;
                }
            }
        }

        private (int, Line) ReadWindowPosition(XmlElement elem)
        {
            int id = elem.GetIntAttribute("id") ?? 0;
            Line position = new Line(TimeBase, TimeBase);

            position.AnchorPoint = GetAnchorPoint(elem.GetIntAttribute("ap") ?? 7);

            int ah = elem.GetIntAttribute("ah") ?? 0;
            int av = elem.GetIntAttribute("av") ?? 0;
            position.Position = GetPixelPosition(new Point(ah, av));
            return (id, position);
        }

        private static (int, Line) ReadWindowStyle(XmlElement elem)
        {
            int id = elem.GetIntAttribute("id") ?? 0;
            Line windowStyle = new Line(TimeBase, TimeBase);

            int printDirection = elem.GetIntAttribute("pd") ?? 0;
            bool scrollDirection = Convert.ToBoolean(elem.GetIntAttribute("sd") ?? 0);
            if (printDirection == 2)
                windowStyle.VerticalTextType = scrollDirection ? VerticalTextType.VerticalLtr : VerticalTextType.VerticalRtl;
            else if (printDirection == 3)
                windowStyle.VerticalTextType = scrollDirection ? VerticalTextType.RotatedRtl : VerticalTextType.RotatedLtr;

            return (id, windowStyle);
        }

        private (int, Section) ReadPen(XmlElement elem)
        {
            int id = elem.GetIntAttribute("id") ?? 0;
            Section pen = new Section();

            int fontStyleId = elem.GetIntAttribute("fs") ?? 0;
            pen.Font = GetFontName(fontStyleId);
            pen.Scale = GetRealFontScale(elem.GetIntAttribute("sz") ?? 100);

            pen.Offset = GetOffsetType(elem.GetIntAttribute("of") ?? 1);
            pen.Bold = Convert.ToBoolean(elem.GetIntAttribute("b") ?? 0);
            pen.Italic = Convert.ToBoolean(elem.GetIntAttribute("i") ?? 0);
            pen.Underline = Convert.ToBoolean(elem.GetIntAttribute("u") ?? 0);

            Color fc = ColorUtil.FromHtml(elem.Attributes["fc"]?.Value ?? "#FFFFFF");
            int fo = elem.GetIntAttribute("fo") ?? 254;
            pen.ForeColor = ColorUtil.ChangeColorAlpha(fc, fo);

            Color bc = ColorUtil.FromHtml(elem.Attributes["bc"]?.Value ?? "#080808");
            int bo = elem.GetIntAttribute("bo") ?? 192;
            pen.BackColor = ColorUtil.ChangeColorAlpha(bc, bo);

            int et = elem.GetIntAttribute("et") ?? 0;
            if (et != 0)
            {
                ShadowType shadowType = GetEdgeType(et);
                Color shadowColor;
                XmlAttribute ecAttr = elem.Attributes["ec"];
                if (ecAttr != null)
                    shadowColor = ColorUtil.ChangeColorAlpha(ColorUtil.FromHtml(ecAttr.Value), 254);
                else
                    shadowColor = Color.FromArgb(fo, 0x22, 0x22, 0x22);

                pen.ShadowColors.Add(shadowType, shadowColor);
            }

            pen.RubyPart = GetRubyPart(elem.GetIntAttribute("rb") ?? 0);
            pen.Packed = Convert.ToBoolean(elem.GetIntAttribute("hg") ?? 0);
            return (id, pen);
        }

        private static Line ReadLine(XmlElement elem, List<Line> positions, List<Line> windowStyles, List<Section> pens)
        {
            int t = elem.GetIntAttribute("t") ?? 0;
            int d = elem.GetIntAttribute("d") ?? 5000;
            if (t == 1) // Reverse Android workaround (See WriteLine)
            {
                t--;
                d++;
            }

            DateTime start = TimeBase.AddMilliseconds(t);
            DateTime end = start.AddMilliseconds(d);
            Line line = new Line(start, end);

            Line position = GetItemAtIndex(positions, elem.GetIntAttribute("wp"));
            if (position != null)
            {
                line.AnchorPoint = position.AnchorPoint;
                line.Position = position.Position;
            }

            Line windowStyle = GetItemAtIndex(windowStyles, elem.GetIntAttribute("ws"));
            if (windowStyle != null)
                line.VerticalTextType = windowStyle.VerticalTextType;

            Section linePen = GetItemAtIndex(pens, elem.GetIntAttribute("p"));
            line.Sections.AddRange(ReadSections(elem, pens, linePen).Where(s => s.Text.Length > 0 || s.StartOffset > TimeSpan.Zero));
            return line;
        }

        private static IEnumerable<Section> ReadSections(XmlElement lineElem, List<Section> pens, Section linePen)
        {
            foreach (XmlNode sectionNode in lineElem.ChildNodes)
            {
                if (sectionNode.NodeType == XmlNodeType.Text)
                    yield return ReadTextSection(sectionNode, linePen);
                else if (sectionNode.NodeType == XmlNodeType.Element)
                    yield return ReadElementSection((XmlElement)sectionNode, pens, linePen);
            }
        }

        private static Section ReadTextSection(XmlNode node, Section linePen)
        {
            Section section = (Section)linePen?.Clone() ?? new Section();
            section.Text = node.Value.Replace(ZeroWidthSpace, "").Replace("\n", "\r\n");
            return section;
        }

        private static Section ReadElementSection(XmlElement elem, List<Section> pens, Section linePen)
        {
            int? penId = elem.GetIntAttribute("p");
            Section pen = GetItemAtIndex(pens, penId) ?? linePen;
            Section section = (Section)pen?.Clone() ?? new Section();
            section.Text = elem.InnerText.Replace(ZeroWidthSpace, "").Replace("\n", "\r\n");

            int? startOffsetMs = elem.GetIntAttribute("t");
            if (startOffsetMs != null)
                section.StartOffset = TimeSpan.FromMilliseconds(startOffsetMs.Value);

            return section;
        }

        private void ApplyEnhancements()
        {
            AddItalicPrefetch();
            for (int i = 0; i < Lines.Count; i++)
            {
                MakeInvisibleTextBlack(i);
                PreventItalicShadowClipping(i);
                PreventBackgroundBoxClipping(i);
                HardenSpaces(i);
                LimitColors(i);
                i += ExpandLineForMultiShadows(i) - 1;
                i += ExpandLineForDarkText(i) - 1;
            }
        }

        /// <summary>
        /// On PC, the first piece of italic text shows up with a noticeable delay: the background appears instantly,
        /// but during a fraction of a second after that, the text is either shown as non-italic or not shown at all.
        /// To work around this, we add an invisible italic subtitle at the start to make YouTube eagerly load
        /// whatever it normally loads lazily.
        /// </summary>
        private void AddItalicPrefetch()
        {
            if (!Lines.SelectMany(l => l.Sections).Any(s => s.Italic))
                return;

            Line italicLine =
                new Line(TimeBase.AddMilliseconds(5000), TimeBase.AddMilliseconds(5100))
                {
                    Position = new PointF(0, 0),
                    AnchorPoint = AnchorPoint.BottomRight
                };
            Section section =
                new Section(ZeroWidthSpace)
                {
                    ForeColor = Color.FromArgb(1, 0, 0, 0),
                    BackColor = Color.Empty,
                    Italic = true
                };
            italicLine.Sections.Add(section);
            Lines.Add(italicLine);
        }

        /// <summary>
        /// The Android app doesn't support text transparency, meaning invisible text would become visible there.
        /// Make such text black so it melts into the app's black, opaque background.
        /// </summary>
        private void MakeInvisibleTextBlack(int lineIndex)
        {
            foreach (Section section in Lines[lineIndex].Sections)
            {
                if (section.ForeColor.A == 0)
                    section.ForeColor = Color.FromArgb(0, 0, 0, 0);
            }
        }

        /// <summary>
        /// Italicized words, such as in "This {\i1}word{\i0} is important", get their shadow cut off
        /// on the right hand side. As a workaround, move the space so we get "This {\i1}word {\i0}is important",
        /// giving the shadow room to fall into.
        /// </summary>
        private void PreventItalicShadowClipping(int lineIndex)
        {
            Line line = Lines[lineIndex];
            for (int i = 0; i < line.Sections.Count - 1; i++)
            {
                Section currSection = line.Sections[i];
                Section nextSection = line.Sections[i + 1];
                if (currSection.Italic && !currSection.Text.EndsWith(" ") && !nextSection.Italic && nextSection.Text.StartsWith(" "))
                {
                    currSection.Text += " ";
                    nextSection.Text = nextSection.Text.Substring(1);
                }
            }
        }

        /// <summary>
        /// Despite several code revisions spanning multiple years, YouTube still doesn't get line breaks in subtitles right.
        /// Even now, having a line break at the beginning or end of a section results in another section losing part of
        /// its background box. The best they could do was apparently getting rid of the rounded corners to make it
        /// less obvious (-> now everything has sharp corners, not just the clipped parts).
        /// The asymmetry in box padding remains visible, however, so we apply a workaround involving zero-width spaces
        /// to fix it where possible.
        /// </summary>
        private void PreventBackgroundBoxClipping(int lineIndex)
        {
            Line line = Lines[lineIndex];
            for (int i = 0; i < line.Sections.Count - 1; i++)
            {
                Section thisSection = line.Sections[i];
                Section nextSection = line.Sections[i + 1];
                if (thisSection.BackColor != nextSection.BackColor ||
                    thisSection.Font != nextSection.Font ||
                    thisSection.Offset != nextSection.Offset ||
                    thisSection.Scale != nextSection.Scale)
                    continue;

                if (thisSection.Text.EndsWith("\r\n"))
                    thisSection.Text = thisSection.Text + ZeroWidthSpace;
                else if (nextSection.Text.StartsWith("\r\n"))
                    nextSection.Text = ZeroWidthSpace + nextSection.Text;
            }
        }

        /// <summary>
        /// Sequences of multiple spaces get collapsed into a single space in browsers -> replace by non-breaking spaces.
        /// (Useful for expanding the background box to cover up on-screen text)
        /// </summary>
        private void HardenSpaces(int lineIndex)
        {
            foreach (Section section in Lines[lineIndex].Sections)
            {
                section.Text = Regex.Replace(section.Text, @"  +", m => new string(' ', m.Value.Length));
            }
        }

        /// <summary>
        /// We should never use 255 for a foreground or background opacity. The reason is that if we do, the upload process
        /// will remove the attribute from the file. For such removed attributes, the player is free to use its own
        /// default settings... which may not match what we want. For example, if the background opacity is lost because
        /// we set it to 255, the PC player will use its default setting of 191. What's more, YouTube allows users to
        /// customize the default settings - in a small window that nobody knows about, with keyboard shortcuts that are
        /// easy to trigger by accident. The result is that some users accidentally configured a low, hard-to-read opacity
        /// and don't know how to change it back.
        /// In short, plenty of reasons to avoid the default opacities.
        /// </summary>
        private void LimitColors(int lineIndex)
        {
            Line line = Lines[lineIndex];
            foreach (Section section in line.Sections)
            {
                if (section.ForeColor.A == 255)
                    section.ForeColor = ColorUtil.ChangeColorAlpha(section.ForeColor, 254);

                if (section.BackColor.A == 255)
                    section.BackColor = ColorUtil.ChangeColorAlpha(section.BackColor, 254);

                // YouTube doesn't have a shadow opacity attribute, so explicitly set the opacity to 255 to avoid creating
                // superfluous <pen>s later on (pens that only differ in shadow opacity). The exception is #222222 which
                // does support opacity, in a way.
                List<ShadowType> shadowTypesToChange = null;
                foreach (KeyValuePair<ShadowType, Color> shadowColor in section.ShadowColors)
                {
                    if (shadowColor.Value.A != 254 && ((shadowColor.Value.ToArgb() & 0xFFFFFF) != 0x222222 || shadowColor.Value.A != section.ForeColor.A))
                        (shadowTypesToChange ??= new List<ShadowType>()).Add(shadowColor.Key);
                }

                if (shadowTypesToChange == null)
                    continue;

                foreach (ShadowType shadowType in shadowTypesToChange)
                {
                    section.ShadowColors[shadowType] = ColorUtil.ChangeColorAlpha(section.ShadowColors[shadowType], 254);
                }
            }
        }

        /// <summary>
        /// YTSubConverter supports multiple shadow types (and colors) on one subtitle by duplicating it as necessary
        /// </summary>
        private int ExpandLineForMultiShadows(int lineIndex)
        {
            Line line = Lines[lineIndex];
            int maxNumShadows = line.Sections.Max(s => s.ShadowColors.Count);
            if (maxNumShadows <= 1)
                return 1;

            List<List<ShadowType>> lineLayerShadowTypes = new List<List<ShadowType>>();
            ShadowType[] orderedShadowTypes = { ShadowType.SoftShadow, ShadowType.HardShadow, ShadowType.Bevel, ShadowType.Glow };
            foreach (Section section in line.Sections)
            {
                List<ShadowType> sectionLayerShadowTypes = new List<ShadowType>();
                foreach (ShadowType shadowType in orderedShadowTypes)
                {
                    if (section.ShadowColors.ContainsKey(shadowType))
                        sectionLayerShadowTypes.Add(shadowType);
                }
                lineLayerShadowTypes.Add(sectionLayerShadowTypes);
            }

            Lines.RemoveAt(lineIndex);

            for (int layerIdx = 0; layerIdx < maxNumShadows; layerIdx++)
            {
                Line shadowLine = (Line)line.Clone();
                for (int sectionIdx = 0; sectionIdx < shadowLine.Sections.Count; sectionIdx++)
                {
                    Section section = shadowLine.Sections[sectionIdx];
                    List<ShadowType> sectionLayerShadowTypes = lineLayerShadowTypes[sectionIdx];

                    if (layerIdx > 0)
                        section.BackColor = Color.Empty;

                    if (layerIdx < sectionLayerShadowTypes.Count)
                        section.ShadowColors.RemoveAll(t => t != sectionLayerShadowTypes[layerIdx]);
                    else
                        section.ShadowColors.Clear();
                }
                Lines.Insert(lineIndex + layerIdx, shadowLine);
            }

            return maxNumShadows;
        }

        /// <summary>
        /// The mobile apps have an unchangeable black background, meaning dark text is unreadable there.
        /// As a workaround, we overlap the dark subtitle with an invisible bright one: this way, we
        /// get the custom background and dark text on PC, and a black background and bright text
        /// on Android (because the Android app doesn't support transparency).
        /// Sadly, this trick doesn't work for iOS: that one supports (only) text transparency,
        /// meaning our bright yet invisible subtitle doesn't show up there.
        /// </summary>
        private int ExpandLineForDarkText(int lineIdx)
        {
            Line line = Lines[lineIdx];
            if (!line.AndroidDarkTextHackAllowed || !line.Sections.Any(s => s.ForeColor.A > 0 && ColorUtil.IsDark(s.ForeColor)))
                return 1;

            Line brightLine = (Line)line.Clone();
            foreach (Section section in brightLine.Sections)
            {
                if (section.ForeColor.A > 0 && ColorUtil.IsDark(section.ForeColor))
                    section.ForeColor = ColorUtil.Brighten(section.ForeColor);

                section.ForeColor = ColorUtil.ChangeColorAlpha(section.ForeColor, 0);
                section.BackColor = ColorUtil.ChangeColorAlpha(section.BackColor, 0);
                section.ShadowColors.Clear();
            }

            Lines.Insert(lineIdx + 1, brightLine);
            return 2;
        }

        /// <summary>
        /// In May of 2020, YouTube introduced a new regression: single-section subtitles that are in the same position
        /// at the same time flicker or are not shown at all on PC. As a workaround, make all single-section subtitles
        /// multi-section by splicing in an empty section.
        /// </summary>
        private void PreventOverlapFlickering()
        {
            List<Line> sortedLines = Lines.ToList();
            sortedLines.Sort((x, y) => x.Start.CompareTo(y.Start));

            for (int i = 0; i < sortedLines.Count; i++)
            {
                Line line1 = sortedLines[i];
                Point line1Position = GetYouTubePosition(line1);

                for (int j = i + 1; j < sortedLines.Count; j++)
                {
                    Line line2 = sortedLines[j];
                    if (line2.Start >= line1.End)
                        break;

                    if (GetYouTubePosition(line2) != line1Position)
                        continue;

                    PreventOverlapFlickering(line1);
                    PreventOverlapFlickering(line2);
                }
            }
        }

        private void PreventOverlapFlickering(Line line)
        {
            if (line.Sections.Count != 1)
                return;

            Section leftSection = line.Sections[0];
            if (string.IsNullOrEmpty(leftSection.Text))
                return;

            // We need to be quite careful when selecting the split point:
            int spliceIndex;
            // If there's taller-than-usual text, we need to make sure both sections get some of it, as otherwise their background boxes won't align
            if ((spliceIndex = leftSection.Text.IndexOf(HasTallBackgroundBox)) >= 0)
                spliceIndex++;
            // Otherwise, we try to splice after a space to prevent cutting off any shadows
            else if ((spliceIndex = leftSection.Text.IndexOf(' ')) >= 0)
                spliceIndex++;
            // If we can't do either of those, we just append the splice section at the end
            else
                spliceIndex = leftSection.Text.Length;

            Section spliceSection = (Section)leftSection.Clone();
            spliceSection.Text = ZeroWidthSpace;
            spliceSection.ForeColor = Color.FromArgb(0, 130, 140, 150);
            line.Sections.Add(spliceSection);

            if (leftSection.Text.Length > spliceIndex)
            {
                Section rightSection = (Section)leftSection.Clone();
                line.Sections.Add(rightSection);

                leftSection.Text = leftSection.Text.Substring(0, spliceIndex);
                rightSection.Text = rightSection.Text.Substring(spliceIndex);
            }

            static bool HasTallBackgroundBox(char c)
            {
                return Util.CharacterRange.HiraganaRange.Contains(c) ||
                       Util.CharacterRange.KatakanaRange.Contains(c) ||
                       Util.CharacterRange.IdeographRange.Contains(c) ||
                       Util.CharacterRange.IdeographExtensionRange.Contains(c) ||
                       Util.CharacterRange.IdeographCompatibilityRange.Contains(c) ||
                       Util.CharacterRange.HangulRange.Contains(c);
            }
        }

        private void WriteHead(XmlWriter writer, Dictionary<Line, int> positions, Dictionary<Line, int> windowStyles, Dictionary<Section, int> pens)
        {
            writer.WriteStartElement("head");

            // The iOS app ignores the background color for the first pen and might have other,
            // similar bugs too, so we write a dummy (unused) item for each of the lists.
            WriteWindowPosition(writer, 0, new Line(TimeBase, TimeBase) { Position = new PointF() });
            foreach (KeyValuePair<Line, int> position in positions)
            {
                WriteWindowPosition(writer, position.Value, position.Key);
            }

            WriteWindowStyle(writer, 0, new Line(TimeBase, TimeBase));
            foreach (KeyValuePair<Line, int> windowStyle in windowStyles)
            {
                WriteWindowStyle(writer, windowStyle.Value, windowStyle.Key);
            }

            WritePen(writer, 0, new Section());
            foreach (KeyValuePair<Section, int> pen in pens)
            {
                WritePen(writer, pen.Value, pen.Key);
            }

            writer.WriteEndElement();
        }

        private void WriteWindowPosition(XmlWriter writer, int positionId, Line position)
        {
            PointF pixelPos = position.Position ?? GetDefaultPosition(position.AnchorPoint);
            Point youtubePos = GetYouTubePosition(pixelPos);

            writer.WriteStartElement("wp");
            writer.WriteAttributeString("id", positionId.ToString());
            writer.WriteAttributeString("ap", GetAnchorPointId(position.AnchorPoint).ToString());
            writer.WriteAttributeString("ah", youtubePos.X.ToString());
            writer.WriteAttributeString("av", youtubePos.Y.ToString());
            writer.WriteEndElement();
        }

        private void WriteWindowStyle(XmlWriter writer, int styleId, Line style)
        {
            writer.WriteStartElement("ws");
            writer.WriteAttributeString("id", styleId.ToString());
            writer.WriteAttributeString("ju", GetJustificationId(style.AnchorPoint).ToString());
            writer.WriteAttributeString("pd", GetTextDirectionId(style.VerticalTextType).ToString());
            writer.WriteAttributeString("sd", IsLineFlowInverted(style.VerticalTextType) ? "1" : "0");
            writer.WriteEndElement();
        }

        private void WritePen(XmlWriter writer, int penId, Section format)
        {
            writer.WriteStartElement("pen");
            writer.WriteAttributeString("id", penId.ToString());

            int fontStyleId = GetFontStyleId(format.Font);
            if (fontStyleId != 0)
                writer.WriteAttributeString("fs", fontStyleId.ToString());
            
            writer.WriteAttributeString("sz", GetYouTubeFontScale(format.Scale).ToString());

            if (format.Offset != OffsetType.Regular)
                writer.WriteAttributeString("of", GetOffsetTypeId(format.Offset).ToString());

            if (format.Bold)
                writer.WriteAttributeString("b", "1");

            if (format.Italic)
                writer.WriteAttributeString("i", "1");

            if (format.Underline)
                writer.WriteAttributeString("u", "1");

            writer.WriteAttributeString("fc", ColorUtil.ToHtml(format.ForeColor));
            writer.WriteAttributeString("fo", format.ForeColor.A.ToString());

            if (format.BackColor.A > 0)
                writer.WriteAttributeString("bc", ColorUtil.ToHtml(format.BackColor));

            writer.WriteAttributeString("bo", format.BackColor.A.ToString());

            if (format.ShadowColors.Count > 0)
            {
                if (format.ShadowColors.Count > 1)
                    throw new NotSupportedException("YTT lines must be reduced to one shadow color before saving");

                KeyValuePair<ShadowType, Color> shadowColor = format.ShadowColors.First();
                if (shadowColor.Value.A > 0)
                {
                    writer.WriteAttributeString("et", GetEdgeTypeId(shadowColor.Key).ToString());

                    // YouTube's handling of shadow transparency is inconsistent: if you specify an "ec" attribute,
                    // the shadow is fully opaque, but if you don't (resulting in a default color of #222222),
                    // it follows the foreground transparency. Because of this, we only write the "ec" attribute
                    // (and lose transparency support) if we have to.
                    if ((shadowColor.Value.ToArgb() & 0xFFFFFF) != 0x222222 ||
                        shadowColor.Value.A != format.ForeColor.A)
                    {
                        writer.WriteAttributeString("ec", ColorUtil.ToHtml(shadowColor.Value));
                    }
                }
            }

            if (format.RubyPart != RubyPart.None)
                writer.WriteAttributeString("rb", GetRubyPartId(format.RubyPart).ToString());

            if (format.Packed)
                writer.WriteAttributeString("hg", "1");

            writer.WriteEndElement();
        }

        private void WriteBody(XmlWriter writer, Dictionary<Line, int> positions, Dictionary<Line, int> windowStyles, Dictionary<Section, int> pens)
        {
            writer.WriteStartElement("body");
            foreach (Line line in Lines)
            {
                WriteLine(writer, line, positions, windowStyles, pens);
            }
            writer.WriteEndElement();
        }

        private void WriteLine(XmlWriter writer, Line line, Dictionary<Line, int> positionIds, Dictionary<Line, int> windowStyleIds, Dictionary<Section, int> penIds)
        {
            if (line.Sections.Count == 0)
                return;

            int lineStartMs = (int)(line.Start - TimeBase).TotalMilliseconds;
            int lineDurationMs = (int)(line.End - line.Start).TotalMilliseconds;

            // If we start in negative time for whatever reason, set the starting time to 1ms and reduce the duration to compensate.
            // (The reason for using 1ms is that the Android app does not respect the positioning of, and sometimes does not display,
            // subtitles that start at 0ms)
            if (lineStartMs <= 0)
            {
                lineDurationMs -= -lineStartMs + 1;
                lineStartMs = 1;
            }

            if (lineDurationMs <= 0)
                return;

            writer.WriteStartElement("p");
            writer.WriteAttributeString("t", lineStartMs.ToString());
            writer.WriteAttributeString("d", lineDurationMs.ToString());
            if (line.Sections.Count == 1)
                writer.WriteAttributeString("p", penIds[line.Sections[0]].ToString());

            writer.WriteAttributeString("wp", positionIds[line].ToString());
            writer.WriteAttributeString("ws", windowStyleIds[line].ToString());

            if (line.Sections.Count == 1)
            {
                writer.WriteValue(AddLineHeightEqualizers(line.Sections[0].Text));
            }
            else
            {
                // The server will remove the "p" (pen ID) attribute of the first section unless the line has text that's not part of any section.
                // We use a zero-width space after the first section to avoid visual impact.
                int multiSectionWorkaroundIdx = line.Sections[0].RubyPart == RubyPart.None ? 0 : 3;
                for (int i = 0; i < line.Sections.Count; i++)
                {
                    WriteSection(writer, line.Sections[i], penIds);
                    if (i == multiSectionWorkaroundIdx)
                        writer.WriteValue(ZeroWidthSpace);
                }
            }

            writer.WriteEndElement();
        }

        private void WriteSection(XmlWriter writer, Section section, Dictionary<Section, int> penIds)
        {
            writer.WriteStartElement("s");
            writer.WriteAttributeString("p", penIds[section].ToString());

            if (section.StartOffset > TimeSpan.Zero)
                writer.WriteAttributeString("t", ((int)section.StartOffset.TotalMilliseconds).ToString());

            writer.WriteValue(AddLineHeightEqualizers(section.Text));
            writer.WriteEndElement();
        }

        // In order to work around YouTube's age-old multisection bug (see WriteLine())
        // we need to add a zero-width space, but doing so changes the line height on
        // some browsers/platforms/fonts. This means we need to make sure *every* line of text
        // has at least one zwsp, as otherwise, subtitles may visibly shift when transitioning
        // from multiple sections (with zwsp) to one section (without zwsp), such as at
        // the last syllable of a karaoke line.
        private static string AddLineHeightEqualizers(string text)
        {
            return Regex.Replace(
                text,
                @"[^\r\n]+",
                m => !m.Value.Contains(ZeroWidthSpace) ? ZeroWidthSpace + m.Value : m.Value
            );
        }

        // YouTube decided to be helpful by moving your subtitles slightly towards the center so they'll never sit at the video's edge.
        // However, it doesn't just impose a cap on each coordinate - it moves your sub regardless of where it is. For example,
        // you don't just get your X = 0% changed to a 2%, but also your 10% to an 11.6%.
        // We counteract this cleverness so our subs actually get displayed where we said they should be.
        // (Or at least as close as possible because, unlike the player, the upload doesn't allow floating point coordinates for whatever reason)
        private Point GetYouTubePosition(Line line)
        {
            return GetYouTubePosition(line.Position ?? GetDefaultPosition(line.AnchorPoint));
        }

        private Point GetYouTubePosition(PointF pixelPos)
        {
            return new Point(
                GetYouTubeCoord(pixelPos.X, VideoDimensions.Width),
                GetYouTubeCoord(pixelPos.Y, VideoDimensions.Height)
            );
        }

        private static int GetYouTubeCoord(float pixelCoord, float maxValue)
        {
            float percentage = pixelCoord / maxValue * 100;
            percentage = (percentage - 2) / 0.96f;
            percentage = Math.Max(percentage, 0);
            percentage = Math.Min(percentage, 100);
            return (int)Math.Round(percentage);
        }

        private PointF GetPixelPosition(Point youtubePos)
        {
            return new PointF(
                GetPixelCoord(youtubePos.X, VideoDimensions.Width),
                GetPixelCoord(youtubePos.Y, VideoDimensions.Height)
            );
        }

        private static float GetPixelCoord(int youtubeCoord, float maxValue)
        {
            return (2 + youtubeCoord * 0.96f) / 100 * maxValue;
        }

        // Similar to positions, YouTube refuses to simply take the specified font scale percentage and apply it.
        // Instead, they do realScale = 1 + (yttScale - 1) / 4, meaning that specifying a 200% scale
        // results in only 125% and that you can't go lower than an actual scale of 75% (yttScale = 0).
        // Maybe they do this to allow for more granularity? But then why not simply allow floating point numbers? Who knows...
        private static float GetRealFontScale(int yttScale)
        {
            float realScale = 1 + ((yttScale / 100f) - 1) / 4;
            return realScale;
        }

        private static int GetYouTubeFontScale(float realScale)
        {
            float yttScale = Math.Max(1 + (realScale - 1) * 4, 0);
            return (int)Math.Round(yttScale * 100);
        }

        private static void AddItemAtIndex<T>(IList<T> list, (int, T) item)
        {
            while (list.Count <= item.Item1)
            {
                list.Add(default);
            }
            list[item.Item1] = item.Item2;
        }

        private static T GetItemAtIndex<T>(IList<T> list, int? index)
        {
            return index != null && index >= 0 && index < list.Count ? list[index.Value] : default;
        }

        private static Dictionary<T, int> ExtractAttributes<T>(IEnumerable<T> objects,  IEqualityComparer<T> comparer)
        {
            Dictionary<T, int> attributes = new Dictionary<T, int>(comparer);
            foreach (T attr in objects)
            {
                if (!attributes.ContainsKey(attr))
                    attributes.Add(attr, 1 + attributes.Count);
            }
            return attributes;
        }

        private static int GetAnchorPointId(AnchorPoint anchorPoint)
        {
            return anchorPoint switch
                   {
                       AnchorPoint.TopLeft => 0,
                       AnchorPoint.TopCenter => 1,
                       AnchorPoint.TopRight => 2,
                       AnchorPoint.MiddleLeft => 3,
                       AnchorPoint.Center => 4,
                       AnchorPoint.MiddleRight => 5,
                       AnchorPoint.BottomLeft => 6,
                       AnchorPoint.BottomCenter => 7,
                       AnchorPoint.BottomRight => 8,
                       _ => throw new ArgumentOutOfRangeException()
                   };
        }

        private static AnchorPoint GetAnchorPoint(int id)
        {
            return id switch
                   {
                       0 => AnchorPoint.TopLeft,
                       1 => AnchorPoint.TopCenter,
                       2 => AnchorPoint.TopRight,
                       3 => AnchorPoint.MiddleLeft,
                       4 => AnchorPoint.Center,
                       5 => AnchorPoint.MiddleRight,
                       6 => AnchorPoint.BottomLeft,
                       7 => AnchorPoint.BottomCenter,
                       8 => AnchorPoint.BottomRight,
                       _ => throw new ArgumentOutOfRangeException()
                   };
        }

        private static int GetJustificationId(AnchorPoint anchorPoint)
        {
            switch (anchorPoint)
            {
                case AnchorPoint.TopLeft:
                case AnchorPoint.MiddleLeft:
                case AnchorPoint.BottomLeft:
                    return 0;

                case AnchorPoint.TopCenter:
                case AnchorPoint.Center:
                case AnchorPoint.BottomCenter:
                    return 2;

                case AnchorPoint.TopRight:
                case AnchorPoint.MiddleRight:
                case AnchorPoint.BottomRight:
                    return 1;

                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static int GetTextDirectionId(VerticalTextType type)
        {
            switch (type)
            {
                case VerticalTextType.VerticalRtl:
                case VerticalTextType.VerticalLtr:
                    return 2;

                case VerticalTextType.RotatedLtr:
                case VerticalTextType.RotatedRtl:
                    return 3;

                default:
                    return 0;
            }
        }

        private static bool IsLineFlowInverted(VerticalTextType type)
        {
            return type == VerticalTextType.VerticalLtr ||
                   type == VerticalTextType.RotatedRtl;
        }

        private static int GetEdgeTypeId(ShadowType type)
        {
            return type switch
                   {
                       ShadowType.HardShadow => 1,
                       ShadowType.Bevel => 2,
                       ShadowType.Glow => 3,
                       ShadowType.SoftShadow => 4,
                       _ => throw new ArgumentOutOfRangeException()
                   };
        }

        private static ShadowType GetEdgeType(int id)
        {
            return id switch
                   {
                       1 => ShadowType.HardShadow,
                       2 => ShadowType.Bevel,
                       3 => ShadowType.Glow,
                       4 => ShadowType.SoftShadow,
                       _ => throw new ArgumentOutOfRangeException()
                   };
        }

        private static int GetOffsetTypeId(OffsetType type)
        {
            return type switch
                   {
                       OffsetType.Subscript => 0,
                       OffsetType.Regular => 1,
                       OffsetType.Superscript => 2,
                       _ => throw new ArgumentOutOfRangeException()
                   };
        }

        private static OffsetType GetOffsetType(int id)
        {
            return id switch
                   {
                       0 => OffsetType.Subscript,
                       1 => OffsetType.Regular,
                       2 => OffsetType.Superscript,
                       _ => throw new ArgumentOutOfRangeException()
                   };
        }

        private static int GetRubyPartId(RubyPart part)
        {
            return part switch
                   {
                       RubyPart.None => 0,
                       RubyPart.Text => 1,
                       RubyPart.Parenthesis => 2,
                       RubyPart.RubyAbove => 4,
                       RubyPart.RubyBelow => 5,
                       _ => throw new ArgumentOutOfRangeException()
                   };
        }

        private static RubyPart GetRubyPart(int id)
        {
            return id switch
                   {
                       0 => RubyPart.None,
                       1 => RubyPart.Text,
                       2 => RubyPart.Parenthesis,
                       4 => RubyPart.RubyAbove,
                       5 => RubyPart.RubyBelow,
                       _ => throw new ArgumentOutOfRangeException()
                   };
        }

        private static int GetFontStyleId(string font)
        {
            switch (font?.ToLower())
            {
                case "courier new":
                case "courier":
                case "nimbus mono l":
                case "cutive mono":
                    return 1;

                case "times new roman":
                case "times":
                case "georgia":
                case "cambria":
                case "pt serif caption":
                    return 2;

                case "deja vu sans mono":       // Support the incorrect font name from YouTube's captions.js,
                case "dejavu sans mono":        // as well as the correct name
                case "lucida console":
                case "monaco":
                case "consolas":
                case "pt mono":
                    return 3;

                case "comic sans ms":
                case "impact":
                case "handlee":
                    return 5;

                case "monotype corsiva":
                case "urw chancery l":
                case "apple chancery":
                case "dancing script":
                    return 6;

                case "carrois gothic sc":
                    return 7;

                default:
                    return 0;
            }
        }

        private static string GetFontName(int fontStyleId)
        {
            return fontStyleId switch
                   {
                       1 => "Courier New",
                       2 => "Times New Roman",
                       3 => "Lucida Console",       // Because of the incorrect "Deja Vu Sans Mono" font name in YouTube's captions.js, browsers fall back to this second option
                       5 => "Comic Sans Ms",
                       6 => "Monotype Corsiva",
                       7 => "Carrois Gothic Sc",
                       _ => "Roboto"
                   };
        }

        private class LinePositionComparer : IEqualityComparer<Line>
        {
            private readonly YttDocument _doc;

            public LinePositionComparer(YttDocument doc)
            {
                _doc = doc;
            }

            public bool Equals(Line x, Line y)
            {
                return x.AnchorPoint == y.AnchorPoint &&
                       _doc.GetYouTubePosition(x) == _doc.GetYouTubePosition(y);
            }

            public int GetHashCode(Line line)
            {
                return line.AnchorPoint.GetHashCode() ^
                       _doc.GetYouTubePosition(line).GetHashCode();
            }
        }

        private class LineAlignmentComparer : IEqualityComparer<Line>
        {
            public bool Equals(Line x, Line y)
            {
                return GetJustificationId(x.AnchorPoint) == GetJustificationId(y.AnchorPoint) &&
                       x.VerticalTextType == y.VerticalTextType;
            }

            public int GetHashCode(Line line)
            {
                return GetJustificationId(line.AnchorPoint) ^
                       line.VerticalTextType.GetHashCode();
            }
        }

        private class NormalizedSectionFormatComparer : SectionFormatComparer
        {
            protected override string NormalizeFont(string font)
            {
                return GetFontName(GetFontStyleId(font));
            }

            protected override float NormalizeScale(float scale)
            {
                return GetRealFontScale(GetYouTubeFontScale(scale));
            }
        }
    }
}
