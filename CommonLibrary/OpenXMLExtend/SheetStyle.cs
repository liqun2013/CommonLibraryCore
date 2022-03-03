using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace CommonLibraryStandard
{
  /// <summary>
  /// 样式
  /// </summary>
  public class StyleItem : IEquatable<StyleItem>
  {
	public StyleItem()
	{
	  FontStyle = new FontStyle();
	  BorderStyle = new BorderStyle();
	  FillStyle = new FillStyle();
	  AlignmentStyle = new AlignmentStyle();
	}
	public FontStyle FontStyle { get; set; }
	public BorderStyle BorderStyle { get; set; }
	public FillStyle FillStyle { get; set; }
	public AlignmentStyle AlignmentStyle { get; set; }
	public bool Equals(StyleItem? o)
	{
	  if (o is null)
		return false;

	  if (ReferenceEquals(this, o))
		return true;

	  return FontStyle.Equals(o.FontStyle) && BorderStyle.Equals(o.BorderStyle) && FillStyle.Equals(o.FillStyle) && AlignmentStyle.Equals(o.AlignmentStyle);
	}

	public bool HasFontStyle =>
	  FontStyle?.FontSize > 0 || FontStyle?.FontBold == true || FontStyle?.FontColor.IsNotNullOrEmptyWhiteSpace() == true || FontStyle?.FontName.IsNotNullOrEmptyWhiteSpace() == true;

	public bool HasFillStyle =>
	  FillStyle.PatternType != Patterns.None || FillStyle.FillColor.IsNotNullOrEmptyWhiteSpace();

	public bool HasBorderStyle =>
	  BorderStyle?.BorderColor.IsNotNullOrEmptyWhiteSpace() == true || BorderStyle?.LeftBorderStyle != BorderStyles.None ||
	  BorderStyle?.RightBorderStyle != BorderStyles.None || BorderStyle?.TopBorderStyle != BorderStyles.None || BorderStyle?.BottomBorderStyle != BorderStyles.None;

	public bool HasAlignmentStyle => AlignmentStyle?.WrapText == true || AlignmentStyle?.VerticalAlignment != VerticalAlignments.Top || AlignmentStyle?.HorizontalAlignment != HorizontalAlignments.General;
	public override int GetHashCode() => FontStyle.GetHashCode() ^ FillStyle.GetHashCode() ^ BorderStyle.GetHashCode() ^ AlignmentStyle.GetHashCode();

	public override bool Equals(object obj) => Equals(obj as StyleItem);
  }

  public class FontStyle : IEquatable<FontStyle>
  {
	public FontStyle()
	{
	  FontSize = 0;
	  FontBold = false;
	  FontColor = "000000";
	  FontName = "Arial";
	}
	public double FontSize { get; set; }
	public bool FontBold { get; set; }
	public string FontName { get; set; }
	public string FontColor { get; set; }
	public bool Equals(FontStyle? o)
	{
	  if (o is null)
		return false;

	  if (ReferenceEquals(this, o))
		return true;

	  return FontSize.Equals(o.FontSize) && FontName.Equals(o.FontName) && FontBold.Equals(o.FontBold) && FontColor.Equals(o.FontColor);
	}

	public override int GetHashCode() => FontSize.GetHashCode() ^ FontName.GetHashCode() ^ FontBold.GetHashCode() ^ FontColor.GetHashCode();

	public override bool Equals(object obj) => Equals(obj as FontStyle);

	public Font ToOpenXmlFont()
	{
	  Font font = new Font();

	  if (FontSize > 0)
		font.AppendChild(new FontSize { Val = new DoubleValue(FontSize) });
	  if (FontBold)
		font.AppendChild(new Bold { Val = FontBold });
	  if (!string.IsNullOrWhiteSpace(FontColor))
		font.AppendChild(new Color { Rgb = FontColor });
	  if (!string.IsNullOrWhiteSpace(FontName))
		font.AppendChild(new FontName { Val = FontName });

	  return font;
	}
  }

  public class FillStyle : IEquatable<FillStyle>
  {
	public FillStyle()
	{
	  PatternType = Patterns.None;
	  FillColor = "FFFFFF";
	}
	public string FillColor { get; set; }
	public Patterns PatternType { get; set; }
	public bool Equals(FillStyle? o)
	{
	  if (o is null)
		return false;

	  if (ReferenceEquals(this, o))
		return true;

	  return FillColor.Equals(o.FillColor) && PatternType.Equals(o.PatternType);
	}

	public override int GetHashCode() => FillColor.GetHashCode() ^ PatternType.GetHashCode();

	public override bool Equals(object obj) => Equals(obj as FillStyle);

	public Fill ToOpenXmlFill()
	{
	  var pf = new PatternFill();
	  if (PatternType != Patterns.None)
		pf.PatternType = ToOpenXmlPatternValue(PatternType);
	  if (FillColor.IsNotNullOrEmptyWhiteSpace())
		pf.ForegroundColor = new ForegroundColor { Rgb = FillColor };

	  var fill = new Fill(pf);
	  return fill;
	}

	private PatternValues ToOpenXmlPatternValue(Patterns p) => (PatternValues)p;
  }

  public class BorderStyle : IEquatable<BorderStyle>
  {
	public BorderStyle()
	{
	  BorderColor = string.Empty;
	  LeftBorderStyle = BorderStyles.Thin;
	  RightBorderStyle = BorderStyles.Thin;
	  TopBorderStyle = BorderStyles.Thin;
	  BottomBorderStyle = BorderStyles.Thin;
	}
	public string BorderColor { get; set; }
	public BorderStyles LeftBorderStyle { get; set; }
	public BorderStyles RightBorderStyle { get; set; }
	public BorderStyles TopBorderStyle { get; set; }
	public BorderStyles BottomBorderStyle { get; set; }
	public bool Equals(BorderStyle? o)
	{
	  if (o is null)
		return false;

	  if (ReferenceEquals(this, o))
		return true;

	  return BorderColor.Equals(o.BorderColor) && LeftBorderStyle.Equals(o.LeftBorderStyle) && TopBorderStyle.Equals(o.TopBorderStyle) &&
		RightBorderStyle.Equals(o.RightBorderStyle) && BottomBorderStyle.Equals(o.BottomBorderStyle);
	}

	public override int GetHashCode() => BorderColor.GetHashCode() ^ LeftBorderStyle.GetHashCode() ^
		TopBorderStyle.GetHashCode() ^ RightBorderStyle.GetHashCode() ^ BottomBorderStyle.GetHashCode();

	public override bool Equals(object obj) => Equals(obj as BorderStyle);

	public Border ToOpenXmlBorder()
	{
	  var brd = new Border();

	  if (TopBorderStyle != BorderStyles.None)
		brd.TopBorder = new TopBorder((BorderColor.IsNotNullOrEmptyWhiteSpace()) ? new Color() { Rgb = BorderColor } : new Color { Auto = true })
		{ Style = ToOpenXmlBorderStyleValue(TopBorderStyle) };
	  if (RightBorderStyle != BorderStyles.None)
		brd.RightBorder = new RightBorder((BorderColor.IsNotNullOrEmptyWhiteSpace()) ? new Color() { Rgb = BorderColor } : new Color { Auto = true })
		{ Style = ToOpenXmlBorderStyleValue(RightBorderStyle) };
	  if (BottomBorderStyle != BorderStyles.None)
		brd.BottomBorder = new BottomBorder((BorderColor.IsNotNullOrEmptyWhiteSpace()) ? new Color() { Rgb = BorderColor } : new Color { Auto = true })
		{ Style = ToOpenXmlBorderStyleValue(BottomBorderStyle) };
	  if (LeftBorderStyle != BorderStyles.None)
		brd.LeftBorder = new LeftBorder((BorderColor.IsNotNullOrEmptyWhiteSpace()) ? new Color() { Rgb = BorderColor } : new Color { Auto = true })
		{ Style = ToOpenXmlBorderStyleValue(LeftBorderStyle) };

	  return brd;
	}
	private BorderStyleValues ToOpenXmlBorderStyleValue(BorderStyles style) => (BorderStyleValues)style;
  }

  public class AlignmentStyle : IEquatable<AlignmentStyle>
  {
	public AlignmentStyle()
	{
	  HorizontalAlignment = HorizontalAlignments.General;
	  VerticalAlignment = VerticalAlignments.Top;
	}
	public bool WrapText { get; set; }
	public HorizontalAlignments HorizontalAlignment { get; set; }
	public VerticalAlignments VerticalAlignment { get; set; }
	public bool Equals(AlignmentStyle? o)
	{
	  if (o is null)
		return false;

	  if (ReferenceEquals(this, o))
		return true;

	  return HorizontalAlignment.Equals(o.HorizontalAlignment) && VerticalAlignment.Equals(o.VerticalAlignment) && WrapText.Equals(o.WrapText);
	}

	public override int GetHashCode() => HorizontalAlignment.GetHashCode() ^ VerticalAlignment.GetHashCode() ^ WrapText.GetHashCode();

	public override bool Equals(object obj) => Equals(obj as AlignmentStyle);

	public Alignment ToOpenXmlAlignment()
	{
	  return new Alignment
	  {
		Horizontal = ToOpenXmlHorizontalAlignmentValues(HorizontalAlignment),
		Vertical = ToOpenXmlVerticalAlignmentValues(VerticalAlignment),
		WrapText = WrapText
	  };
	}

	private VerticalAlignmentValues ToOpenXmlVerticalAlignmentValues(VerticalAlignments style) => (VerticalAlignmentValues)style;
	private HorizontalAlignmentValues ToOpenXmlHorizontalAlignmentValues(HorizontalAlignments style) => (HorizontalAlignmentValues)style;
  }

  #region enums
  public enum HorizontalAlignments
  {
	[EnumString("general")]
	General,
	[EnumString("left")]
	Left,
	[EnumString("center")]
	Center,
	[EnumString("right")]
	Right,
	[EnumString("fill")]
	Fill,
	[EnumString("justify")]
	Justify,
	[EnumString("centerContinuous")]
	CenterContinuous,
	[EnumString("distributed")]
	Distributed
  }
  public enum VerticalAlignments
  {
	[EnumString("top")]
	Top,
	[EnumString("center")]
	Center,
	[EnumString("bottom")]
	Bottom,
	[EnumString("justify")]
	Justify,
	[EnumString("distributed")]
	Distributed
  }

  /// <summary>
  /// Border Line Styles
  /// </summary>
  public enum BorderStyles
  {
	[EnumString("none")]
	None,
	[EnumString("thin")]
	Thin,
	[EnumString("medium")]
	Medium,
	[EnumString("dashed")]
	Dashed,
	[EnumString("dotted")]
	Dotted,
	[EnumString("thick")]
	Thick,
	[EnumString("double")]
	Double,
	[EnumString("hair")]
	Hair,
	[EnumString("mediumDashed")]
	MediumDashed,
	[EnumString("dashDot")]
	DashDot,
	[EnumString("mediumDashDot")]
	MediumDashDot,
	[EnumString("dashDotDot")]
	DashDotDot,
	[EnumString("mediumDashDotDot")]
	MediumDashDotDot,
	[EnumString("slantDashDot")]
	SlantDashDot
  }

  /// <summary>
  /// Pattern Type
  /// </summary>
  public enum Patterns
  {
	[EnumString("none")]
	None,
	[EnumString("solid")]
	Solid,
	[EnumString("mediumGray")]
	MediumGray,
	[EnumString("darkGray")]
	DarkGray,
	[EnumString("lightGray")]
	LightGray,
	[EnumString("darkHorizontal")]
	DarkHorizontal,
	[EnumString("darkVertical")]
	DarkVertical,
	[EnumString("darkDown")]
	DarkDown,
	[EnumString("darkUp")]
	DarkUp,
	[EnumString("darkGrid")]
	DarkGrid,
	[EnumString("darkTrellis")]
	DarkTrellis,
	[EnumString("lightHorizontal")]
	LightHorizontal,
	[EnumString("lightVertical")]
	LightVertical,
	[EnumString("lightDown")]
	LightDown,
	[EnumString("lightUp")]
	LightUp,
	[EnumString("lightGrid")]
	LightGrid,
	[EnumString("lightTrellis")]
	LightTrellis,
	[EnumString("gray125")]
	Gray125,
	[EnumString("gray0625")]
	Gray0625
  }
  #endregion
}