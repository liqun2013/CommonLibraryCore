namespace CommonLibraryCore;
public static class TypeExtenssions
{
  #region string convert extenssions

  public static T ToEnum<T>(this string s, bool ignoreCase, bool useDefaultWhenExcpOccurs, T defaultValue = default) where T : struct
  {
	T result = defaultValue;
	if (!string.IsNullOrWhiteSpace(s))
	{
	  if (Enum.TryParse(s, ignoreCase, out T v))
		result = v;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	}
	else
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }
  public static char ToChar(this string s, bool useDefaultWhenExcpOccurs, char defaultValue = ' ')
  {
	char result = defaultValue;

	if (!string.IsNullOrEmpty(s))
	{
	  if (char.TryParse(s.Trim(), out char c))
		result = c;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static char? ToNullableChar(this string s, bool useDefaultWhenExcpOccurs, char? defaultValue = null)
  {
	char? result = defaultValue;

	if (s != null)
	{
	  if (char.TryParse(s.Trim(), out char c))
		result = c;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	}

	return result;
  }
  public static bool ToBoolean(this string s, bool useDefaultWhenExcpOccurs, bool defaultValue = false)
  {
	bool result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  if (s.Trim().Equals("1"))
		result = true;
	  else if (s.Trim().Equals("0"))
		result = false;
	  else if (bool.TryParse(s.Trim(), out bool v))
		result = v;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static bool? ToNullableBoolean(this string s, bool useDefaultWhenExcpOccurs, bool? defaultValue = null)
  {
	bool? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  if (s.Trim().Equals("1"))
		result = true;
	  else if (s.Trim().Equals("0"))
		result = false;
	  else if (bool.TryParse(s.Trim(), out bool v))
		result = v;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	}

	return result;
  }
  public static short ToInt16(this string s, bool useDefaultWhenExcpOccurs, short defaultValue = 0)
  {
	short result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  short v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (short.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (short.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static short? ToNullableInt16(this string s, bool useDefaultWhenExcpOccurs, short? defaultValue = null)
  {
	short? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  short v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (short.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (short.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}

	return result;
  }

  public static ushort ToUInt16(this string s, bool useDefaultWhenExcpOccurs, ushort defaultValue = 0)
  {
	ushort result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  ushort v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (ushort.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (ushort.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static ushort? ToNullableUInt16(this string s, bool useDefaultWhenExcpOccurs, ushort? defaultValue = null)
  {
	ushort? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  ushort v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (ushort.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (ushort.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}

	return result;
  }

  public static byte ToByte(this string s, bool useDefaultWhenExcpOccurs, byte defaultValue = 0)
  {
	byte result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  if (byte.TryParse(s.Trim(), out byte v))
		result = v;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static byte? ToNullableByte(this string s, bool useDefaultWhenExcpOccurs, byte? defaultValue = null)
  {
	byte? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  if (byte.TryParse(s.Trim(), out byte v))
		result = v;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	}

	return result;
  }

  public static sbyte ToSbyte(this string s, bool useDefaultWhenExcpOccurs, sbyte defaultValue = 0)
  {
	sbyte result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  if (sbyte.TryParse(s.Trim(), out sbyte v))
		result = v;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static sbyte? ToNullableSbyte(this string s, bool useDefaultWhenExcpOccurs, sbyte? defaultValue = null)
  {
	sbyte? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  if (sbyte.TryParse(s.Trim(), out sbyte v))
		result = v;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	}

	return result;
  }

  public static int ToInt32(this string s, bool useDefaultWhenExcpOccurs, int defaultValue = 0)
  {
	int result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  int v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (int.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (int.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static uint ToUInt32(this string s, bool useDefaultWhenExcpOccurs, uint defaultValue = 0)
  {
	uint result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  uint v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (uint.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (uint.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static int? ToNullableInt32(this string s, bool useDefaultWhenExcpOccurs, int? defaultValue = null)
  {
	int? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  int v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (int.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (int.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}

	return result;
  }

  public static uint? ToNullableUInt32(this string s, bool useDefaultWhenExcpOccurs, uint? defaultValue = null)
  {
	uint? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  uint v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (uint.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (uint.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}

	return result;
  }

  public static long ToInt64(this string s, bool useDefaultWhenExcpOccurs, long defaultValue = 0)
  {
	long result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  long v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (long.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (long.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static long? ToNullableInt64(this string s, bool useDefaultWhenExcpOccurs, long? defaultValue = null)
  {
	long? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  long v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (long.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (long.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}

	return result;
  }

  public static ulong ToUInt64(this string s, bool useDefaultWhenExcpOccurs, ulong defaultValue = 0)
  {
	ulong result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  ulong v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (ulong.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (ulong.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static ulong? ToNullableUInt64(this string s, bool useDefaultWhenExcpOccurs, ulong? defaultValue = null)
  {
	ulong? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  ulong v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (ulong.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (ulong.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}

	return result;
  }

  public static float ToFloat(this string s, bool useDefaultWhenExcpOccurs, float defaultValue = 0)
  {
	float result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  float v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (float.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (float.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static float? ToNullableFloat(this string s, bool useDefaultWhenExcpOccurs, float? defaultValue = null)
  {
	float? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  float v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (float.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (float.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}

	return result;
  }
  public static decimal ToDecimal(this string s, bool useDefaultWhenExcpOccurs, decimal defaultValue = 0)
  {
	decimal result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  decimal v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (decimal.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (decimal.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static decimal? ToNullableDecimal(this string s, bool useDefaultWhenExcpOccurs, decimal? defaultValue = null)
  {
	decimal? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  decimal v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (decimal.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (decimal.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}

	return result;
  }

  public static double ToDouble(this string s, bool useDefaultWhenExcpOccurs, double defaultValue = 0)
  {
	double result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  double v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (double.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (double.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static double? ToNullableDouble(this string s, bool useDefaultWhenExcpOccurs, double? defaultValue = null)
  {
	double? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  double v;
	  if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
	  {
		if (double.TryParse(s.Trim(), NumberStyles.Float, null, out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	  else
	  {
		if (double.TryParse(s.Trim(), out v))
		  result = v;
		else if (!useDefaultWhenExcpOccurs)
		  throw new FormatException($"The parameter {nameof(s)}:{s} format could not be converted.");
	  }
	}

	return result;
  }

  public static DateTime ToDateTime(this string s, string format, bool useDefaultWhenExcpOccurs, DateTime defaultValue = new DateTime())
  {
	DateTime result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  if (!string.IsNullOrWhiteSpace(format) && DateTime.TryParseExact(s.Trim(), format, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime v))
		result = v;
	  else if (DateTime.TryParse(s.Trim(), out v))
		result = v;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s}/format:{format} could not be converted.");
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static DateTime? ToNullableDateTime(this string s, string format, bool useDefaultWhenExcpOccurs, DateTime? defaultValue = null)
  {
	DateTime? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  if (DateTime.TryParseExact(s.Trim(), format, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime v))
		result = v;
	  else if (DateTime.TryParse(s.Trim(), out v))
		result = v;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s}/format:{format} could not be converted.");
	}

	return result;
  }
  public static DateTime ToDateTime(this string s, bool useDefaultWhenExcpOccurs, DateTime defaultValue = new DateTime())
  {
	DateTime result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  if (DateTime.TryParse(s.Trim(), out DateTime v))
		result = v;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s} could not be converted.");
	}
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  public static DateTime? ToNullableDateTime(this string s, bool useDefaultWhenExcpOccurs, DateTime? defaultValue = null)
  {
	DateTime? result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	{
	  if (DateTime.TryParse(s.Trim(), out DateTime v))
		result = v;
	  else if (!useDefaultWhenExcpOccurs)
		throw new FormatException($"The parameter {nameof(s)}:{s} could not be converted.");
	}

	return result;
  }

  /// <summary>
  /// 从此实例检索子字符串。 子字符串从指定的字符位置开始且具有指定的长度。
  /// </summary>
  /// <param name="start">起始字符位置（从零开始）。</param>
  /// <param name="length">子字符串中的字符数。(当length大于字符串长度时取字符串长度) </param>
  /// <param name="useDefaultWhenExcpOccurs">出现异常时是否使用默认值</param>
  /// <returns></returns>
  public static string SubString(this string s, int start, int length, bool useDefaultWhenExcpOccurs, string defaultValue = "")
  {
	string result = defaultValue;

	if (!string.IsNullOrWhiteSpace(s))
	  result = s.Substring(start, Math.Min(s.Length - start, length));
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(s), $"The parameter {nameof(s)} is null or empty.");

	return result;
  }

  #endregion

  #region object convert extenssions
  public static bool ToBoolean(this object o, bool useDefaultWhenExcpOccurs, bool defaultValue = false)
  {
	bool result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToBoolean(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static bool? ToNullableBoolean(this object o, bool useDefaultWhenExcpOccurs, bool? defaultValue = null)
  {
	bool? result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableBoolean(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }
  public static short ToInt16(this object o, bool useDefaultWhenExcpOccurs, short defaultValue = 0)
  {
	short result = defaultValue;
	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToInt16(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static short? ToNullableInt16(this object o, bool useDefaultWhenExcpOccurs, short? defaultValue = null)
  {
	short? result = defaultValue;
	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableInt16(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }

  public static ushort ToUInt16(this object o, bool useDefaultWhenExcpOccurs, ushort defaultValue = 0)
  {
	ushort result = defaultValue;
	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToUInt16(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static ushort? ToNullableUInt16(this object o, bool useDefaultWhenExcpOccurs, ushort? defaultValue = null)
  {
	ushort? result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableUInt16(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }

  public static int ToInt32(this object o, bool useDefaultWhenExcpOccurs, int defaultValue = 0)
  {
	int result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToInt32(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static uint ToUInt32(this object o, bool useDefaultWhenExcpOccurs, uint defaultValue = 0)
  {
	uint result = defaultValue;
	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToUInt32(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static int? ToNullableInt32(this object o, bool useDefaultWhenExcpOccurs, int? defaultValue = null)
  {
	int? result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableInt32(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }

  public static uint? ToNullableUInt32(this object o, bool useDefaultWhenExcpOccurs, uint? defaultValue = null)
  {
	uint? result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableUInt32(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }

  public static long ToInt64(this object o, bool useDefaultWhenExcpOccurs, long defaultValue = 0)
  {
	long result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToInt64(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static long? ToNullableInt64(this object o, bool useDefaultWhenExcpOccurs, long? defaultValue = null)
  {
	long? result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableInt64(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }

  public static ulong ToUInt64(this object o, bool useDefaultWhenExcpOccurs, ulong defaultValue = 0)
  {
	ulong result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToUInt64(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static ulong? ToNullableUInt64(this object o, bool useDefaultWhenExcpOccurs, ulong? defaultValue = null)
  {
	ulong? result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableUInt64(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }

  public static float ToFloat(this object o, bool useDefaultWhenExcpOccurs, float defaultValue = 0)
  {
	float result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToFloat(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static float? ToNullableFloat(this object o, bool useDefaultWhenExcpOccurs, float? defaultValue = null)
  {
	float? result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableFloat(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }
  public static decimal ToDecimal(this object o, bool useDefaultWhenExcpOccurs, decimal defaultValue = 0)
  {
	decimal result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToDecimal(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static decimal? ToNullableDecimal(this object o, bool useDefaultWhenExcpOccurs, decimal? defaultValue = null)
  {
	decimal? result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableDecimal(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }

  public static double ToDouble(this object o, bool useDefaultWhenExcpOccurs, double defaultValue = 0)
  {
	double result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToDouble(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static double? ToNullableDouble(this object o, bool useDefaultWhenExcpOccurs, double? defaultValue = null)
  {
	double? result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableDouble(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }

  public static DateTime ToDateTime(this object o, string format, bool useDefaultWhenExcpOccurs, DateTime defaultValue = default)
  {
	DateTime result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToDateTime(format, useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }
  public static DateTime ToDateTime(this object o, bool useDefaultWhenExcpOccurs, DateTime defaultValue = default)
  {
	DateTime result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToDateTime(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static DateTime? ToNullableDateTime(this object o, string format, bool useDefaultWhenExcpOccurs, DateTime? defaultValue = null)
  {
	DateTime? result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableDateTime(format, useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }
  public static DateTime? ToNullableDateTime(this object o, bool useDefaultWhenExcpOccurs, DateTime? defaultValue = null)
  {
	DateTime? result = null;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableDateTime(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }

  public static char ToChar(this object o, bool useDefaultWhenExcpOccurs, char defaultValue = ' ')
  {
	char result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToChar(useDefaultWhenExcpOccurs, defaultValue);
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }

  public static char? ToNullableChar(this object o, bool useDefaultWhenExcpOccurs, char? defaultValue = null)
  {
	char? result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString().ToNullableChar(useDefaultWhenExcpOccurs, defaultValue);

	return result;
  }
  public static string ToString(this object o, bool useDefaultWhenExcpOccurs, string defaultValue = default)
  {
	string result = defaultValue;

	if (o != null && o != DBNull.Value)
	  result = o.ToString();
	else if (!useDefaultWhenExcpOccurs)
	  throw new ArgumentNullException(nameof(o), $"The parameter {nameof(o)} is null or empty.");

	return result;
  }
  #endregion

  #region convert number to english words
  public static string ToEnglishWords(this int n)
  {
	return new EnglishNumberToWordsConverter().Convert(n);
  }
  public static string ToEnglishWords(this long n)
  {
	return new EnglishNumberToWordsConverter().Convert(n);
  }
  #endregion

  public static bool IsNumeric(this string s)
  {
	return double.TryParse(s, NumberStyles.Any, NumberFormatInfo.InvariantInfo, out _);
  }

  public static bool IsEmailAddress(this string s)
  {
	string emailRegex = @"^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$";
	if (string.IsNullOrWhiteSpace(s))
	  return false;

	var regex = new Regex(emailRegex);
	return regex.IsMatch(s);
  }

  public static bool IsNumeric(this object o)
  {
	bool result = false;
	if (o != null && o != DBNull.Value)
	{
	  Type objType = o.GetType();
	  objType = Nullable.GetUnderlyingType(objType) ?? objType;

	  if (objType.Name.Equals("string", StringComparison.OrdinalIgnoreCase))
		result = double.TryParse(o.ToString(), NumberStyles.Any, NumberFormatInfo.InvariantInfo, out _);
	  else
		result = objType == typeof(sbyte) || objType == typeof(byte) || objType == typeof(short) ||
		  objType == typeof(ushort) || objType == typeof(int) || objType == typeof(uint) || objType == typeof(long) ||
		  objType == typeof(ulong) || objType == typeof(float) || objType == typeof(double) || objType == typeof(decimal) || objType == typeof(System.Numerics.BigInteger);
	}
	return result;
  }

  #region preconditions
  public static T CheckNotNull<T>(this T obj)
  {
	if (obj is null)
	{
	  ArgumentNullException argumentNullException = new();
	  throw argumentNullException;
	}
	else
	  return obj;
  }
  public static object CheckNotDBNull(this object obj)
  {
	if (obj == DBNull.Value || obj is null)
	{
	  ArgumentNullException argumentNullException = new();
	  throw argumentNullException;
	}
	else
	  return obj;
  }
  public static bool IsNotNullOrEmptyWhiteSpace(this string s) => !string.IsNullOrWhiteSpace(s);
  #endregion

  #region openxmlextend
  /// <summary>
  /// 转换成openxml格式的数据，用于导入excel
  /// </summary>
  /// <typeparam name="T"></typeparam>
  /// <param name="lstData"></param>
  /// <param name="addHeadRow"></param>
  /// <returns></returns>
  public static ExlsSheetData ToExlsSheetData<T>(this IEnumerable<T> lstData, bool addHeadRow) where T : new()
  {
	return new ExlsSheetData().ToSheetdata(lstData, addHeadRow);
  }
  #endregion
}
