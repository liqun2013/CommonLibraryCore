namespace CommonLibraryCore
{
  public static class TypeExtenssions
	{
		#region string convert extenssions

		public static T ToEnum<T>(this string s, bool ignoreCase, bool useDefaultWhenExcpOccurs, T defaultValue = default(T)) where T : struct
		{
			T result = defaultValue;
			if (!string.IsNullOrWhiteSpace(s))
			{
				T v = default(T);
				if (Enum.TryParse(s, ignoreCase, out v))
					result = v;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
			}
			else
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}
		public static char ToChar(this string s, bool useDefaultWhenExcpOccurs, char defaultValue = ' ')
		{
			char result = defaultValue;

			if (!string.IsNullOrEmpty(s))
			{
				char c = default(char);
				if (char.TryParse(s.Trim(), out c))
					result = c;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static char? ToNullableChar(this string s, bool useDefaultWhenExcpOccurs, char? defaultValue = null)
		{
			char? result = defaultValue;

			if (s != null)
			{
				char c = default(char);
				if (char.TryParse(s.Trim(), out c))
					result = c;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
			}

			return result;
		}
		public static bool ToBoolean(this string s, bool useDefaultWhenExcpOccurs, bool defaultValue = false)
		{
			bool result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				bool v = default(bool);
				if (s.Trim().Equals("1"))
					result = true;
				else if (s.Trim().Equals("0"))
					result = false;
				else if (bool.TryParse(s.Trim(), out v))
					result = v;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static bool? ToNullableBoolean(this string s, bool useDefaultWhenExcpOccurs, bool? defaultValue = null)
		{
			bool? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				bool v = default(bool);
				if (s.Trim().Equals("1"))
					result = true;
				else if (s.Trim().Equals("0"))
					result = false;
				else if (bool.TryParse(s.Trim(), out v))
					result = v;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
			}

			return result;
		}
		public static short ToInt16(this string s, bool useDefaultWhenExcpOccurs, short defaultValue = 0)
		{
			short result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				short v = default(short);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (short.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (short.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static short? ToNullableInt16(this string s, bool useDefaultWhenExcpOccurs, short? defaultValue = null)
		{
			short? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				short v = default(short);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (short.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (short.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}

			return result;
		}

		public static ushort ToUInt16(this string s, bool useDefaultWhenExcpOccurs, ushort defaultValue = 0)
		{
			ushort result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				ushort v = default(ushort);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (ushort.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (ushort.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static ushort? ToNullableUInt16(this string s, bool useDefaultWhenExcpOccurs, ushort? defaultValue = null)
		{
			ushort? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				ushort v = default(ushort);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (ushort.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (ushort.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}

			return result;
		}

		public static byte ToByte(this string s, bool useDefaultWhenExcpOccurs, byte defaultValue = 0)
		{
			byte result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				byte v = default(byte);

				if (byte.TryParse(s.Trim(), out v))
					result = v;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static byte? ToNullableByte(this string s, bool useDefaultWhenExcpOccurs, byte? defaultValue = null)
		{
			byte? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				byte v = default(byte);
				if (byte.TryParse(s.Trim(), out v))
					result = v;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
			}

			return result;
		}

		public static sbyte ToSbyte(this string s, bool useDefaultWhenExcpOccurs, sbyte defaultValue = 0)
		{
			sbyte result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				sbyte v = default(sbyte);
				if (sbyte.TryParse(s.Trim(), out v))
					result = v;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static sbyte? ToNullableSbyte(this string s, bool useDefaultWhenExcpOccurs, sbyte? defaultValue = null)
		{
			sbyte? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				sbyte v = default(sbyte);
				if (sbyte.TryParse(s.Trim(), out v))
					result = v;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
			}

			return result;
		}

		public static int ToInt32(this string s, bool useDefaultWhenExcpOccurs, int defaultValue = 0)
		{
			int result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				int v = default(int);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (int.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (int.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static uint ToUInt32(this string s, bool useDefaultWhenExcpOccurs, uint defaultValue = 0)
		{
			uint result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				uint v = default(uint);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (uint.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (uint.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static int? ToNullableInt32(this string s, bool useDefaultWhenExcpOccurs, int? defaultValue = null)
		{
			int? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				int v = default(int);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (int.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (int.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}

			return result;
		}

		public static uint? ToNullableUInt32(this string s, bool useDefaultWhenExcpOccurs, uint? defaultValue = null)
		{
			uint? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				uint v = default(uint);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (uint.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (uint.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}

			return result;
		}

		public static long ToInt64(this string s, bool useDefaultWhenExcpOccurs, long defaultValue = 0)
		{
			long result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				long v = default(long);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (long.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (long.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static long? ToNullableInt64(this string s, bool useDefaultWhenExcpOccurs, long? defaultValue = null)
		{
			long? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				long v = default(long);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (long.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (long.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}

			return result;
		}

		public static ulong ToUInt64(this string s, bool useDefaultWhenExcpOccurs, ulong defaultValue = 0)
		{
			ulong result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				ulong v = default(ulong);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (ulong.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (ulong.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static ulong? ToNullableUInt64(this string s, bool useDefaultWhenExcpOccurs, ulong? defaultValue = null)
		{
			ulong? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				ulong v = default(ulong);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (ulong.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (ulong.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}

			return result;
		}

		public static float ToFloat(this string s, bool useDefaultWhenExcpOccurs, float defaultValue = 0)
		{
			float result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				float v = default(float);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (float.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (float.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static float? ToNullableFloat(this string s, bool useDefaultWhenExcpOccurs, float? defaultValue = null)
		{
			float? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				float v = default(float);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (float.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (float.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}

			return result;
		}
		public static decimal ToDecimal(this string s, bool useDefaultWhenExcpOccurs, decimal defaultValue = 0)
		{
			decimal result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				decimal v = default(decimal);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (decimal.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (decimal.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static decimal? ToNullableDecimal(this string s, bool useDefaultWhenExcpOccurs, decimal? defaultValue = null)
		{
			decimal? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				decimal v = default(decimal);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (decimal.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (decimal.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}

			return result;
		}

		public static double ToDouble(this string s, bool useDefaultWhenExcpOccurs, double defaultValue = 0)
		{
			double result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				double v = default(double);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (double.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (double.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static double? ToNullableDouble(this string s, bool useDefaultWhenExcpOccurs, double? defaultValue = null)
		{
			double? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				double v = default(double);
				if (s.Contains("E-") || s.Contains("e-") || s.Contains("E+") || s.Contains("e+"))
				{
					if (double.TryParse(s.Trim(), NumberStyles.Float, null, out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
				else
				{
					if (double.TryParse(s.Trim(), out v))
						result = v;
					else if (!useDefaultWhenExcpOccurs)
						throw new FormatException(string.Format("The parameter s:{0} format could not be converted.", s));
				}
			}

			return result;
		}

		public static DateTime ToDateTime(this string s, string format, bool useDefaultWhenExcpOccurs, DateTime defaultValue = new DateTime())
		{
			DateTime result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				DateTime v = default(DateTime);
				if (!string.IsNullOrWhiteSpace(format) && DateTime.TryParseExact(s.Trim(), format, CultureInfo.CurrentCulture, DateTimeStyles.None, out v))
					result = v;
				else if (DateTime.TryParse(s.Trim(), out v))
					result = v;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0}/format:{1} could not be converted.", s, format));
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static DateTime? ToNullableDateTime(this string s, string format, bool useDefaultWhenExcpOccurs, DateTime? defaultValue = null)
		{
			DateTime? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				DateTime v = default(DateTime);
				if (DateTime.TryParseExact(s.Trim(), format, CultureInfo.CurrentCulture, DateTimeStyles.None, out v))
					result = v;
				else if (DateTime.TryParse(s.Trim(), out v))
					result = v;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0}/format:{1} could not be converted.", s, format));
			}

			return result;
		}
		public static DateTime ToDateTime(this string s, bool useDefaultWhenExcpOccurs, DateTime defaultValue = new DateTime())
		{
			DateTime result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				DateTime v = default(DateTime);
				if (DateTime.TryParse(s.Trim(), out v))
					result = v;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0} could not be converted.", s));
			}
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

			return result;
		}

		public static DateTime? ToNullableDateTime(this string s, bool useDefaultWhenExcpOccurs, DateTime? defaultValue = null)
		{
			DateTime? result = defaultValue;

			if (!string.IsNullOrWhiteSpace(s))
			{
				DateTime v = default(DateTime);
				if (DateTime.TryParse(s.Trim(), out v))
					result = v;
				else if (!useDefaultWhenExcpOccurs)
					throw new FormatException(string.Format("The parameter s:{0} could not be converted.", s));
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
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

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
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

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
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

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
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

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
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

			return result;
		}

		public static uint ToUInt32(this object o, bool useDefaultWhenExcpOccurs, uint defaultValue = 0)
		{
			uint result = defaultValue;
			if (o != null && o != DBNull.Value)
				result = o.ToString().ToUInt32(useDefaultWhenExcpOccurs, defaultValue);
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

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
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

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
				throw new ArgumentNullException("s", "The parameter s is null or empty.");

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
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

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
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

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
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

			return result;
		}

		public static double? ToNullableDouble(this object o, bool useDefaultWhenExcpOccurs, double? defaultValue = null)
		{
			double? result = defaultValue;

			if (o != null && o != DBNull.Value)
				result = o.ToString().ToNullableDouble(useDefaultWhenExcpOccurs, defaultValue);

			return result;
		}

		public static DateTime ToDateTime(this object o, string format, bool useDefaultWhenExcpOccurs, DateTime defaultValue = default(DateTime))
		{
			DateTime result = defaultValue;

			if (o != null && o != DBNull.Value)
				result = o.ToString().ToDateTime(useDefaultWhenExcpOccurs, defaultValue);
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

			return result;
		}
		public static DateTime ToDateTime(this object o, bool useDefaultWhenExcpOccurs, DateTime defaultValue = default(DateTime))
		{
			DateTime result = defaultValue;

			if (o != null && o != DBNull.Value)
				result = o.ToString().ToDateTime(useDefaultWhenExcpOccurs, defaultValue);
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

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
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

			return result;
		}

		public static char? ToNullableChar(this object o, bool useDefaultWhenExcpOccurs, char? defaultValue = null)
		{
			char? result = defaultValue;

			if (o != null && o != DBNull.Value)
				result = o.ToString().ToNullableChar(useDefaultWhenExcpOccurs, defaultValue);

			return result;
		}
		public static string ToString(this object o, bool useDefaultWhenExcpOccurs, string defaultValue = default(string))
		{
			string result = defaultValue;

			if (o != null && o != DBNull.Value)
				result = o.ToString();
			else if (!useDefaultWhenExcpOccurs)
				throw new ArgumentNullException("o", "The parameter o is null or empty.");

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
			double d = 0;
			return double.TryParse(s, NumberStyles.Any, NumberFormatInfo.InvariantInfo, out d);
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

				if (objType.IsPrimitive)
				{
					result = objType == typeof(sbyte) || objType == typeof(byte) || objType == typeof(short) || objType == typeof(ushort) || objType == typeof(int)
						|| objType == typeof(uint) || objType == typeof(long) || objType == typeof(ulong) || objType == typeof(float) || objType == typeof(double) || objType == typeof(decimal);
				}
				else if (objType.Name.Equals("string", StringComparison.OrdinalIgnoreCase))
				{
					double d = 0;
					result = double.TryParse(o.ToString(), NumberStyles.Any, NumberFormatInfo.InvariantInfo, out d);
				}
			}
			return result;
		}

		#region preconditions
		public static T CheckNotNull<T>(this T obj)
		{
			if (obj == null)
				throw new ArgumentNullException();
			else
				return obj;
		}
		public static object CheckNotDBNull(this object obj)
		{
			if (obj == DBNull.Value || obj == null)
				throw new ArgumentNullException();
			else
				return obj;
		}
		#endregion

	}
}
