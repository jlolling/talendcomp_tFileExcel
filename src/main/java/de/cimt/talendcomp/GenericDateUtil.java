package de.cimt.talendcomp;



import java.lang.reflect.Field;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Util-Klasse zum generischen Extrahieren eines Datums aus einem Datumsstring.
 * Im Gegensatz zu anderen Klassen die diese Aufgabe &uuml;bernehmen erfolgt die
 * Extraktion automatisch, der Entwickler ist daher nicht gezwungen das
 * Datumsformat mit anzugeben. Die Klasse enth&auml;lt eine Konfiguration
 * mehrerer unterst&uuml;tzter Datumsformate, die unterst&uuml;tzten Formate
 * k&ouml;nnen &uuml;ber die Methode {@link #getSupportedFormats()} abgefragt
 * werden. Sollten weitere Datumsformate ben&ouml;tigt werden, muss einfach ein
 * neuer Konfigurationssatz hinzugef&uuml;gt werden, an den Methoden muss keine
 * &Auml;nderung erfolgen.
 * 
 * @author x65721 - André Hermann
 * @version 1.0
 * @since 17.07.2013
 */
public class GenericDateUtil {

	/**
	 * Innere Klasse um die Verwendung des Singleton-Pattern zu
	 * erm&ouml;glichen, da bei statischen Zugriffen viel zuviele Exceptions
	 * durchgereicht werden m&uuml;ssen und diese Util-Klasse per Class.forName
	 * ermittelt werden muss, was das ganz fehleranf&auml;lliger macht.
	 * 
	 * @author André Hermann
	 * @version 1.0
	 * @since 17.07.2013
	 */
	protected static class DateUtil {

		/**
		 * Die Klasse ist nach dem Singleton-Pattern implementiert und wird
		 * threadsave initialisiert.
		 */
		protected static DateUtil instance = new DateUtil();

		// Textfragmente
		protected final static String US_DAYS_E = "Mon|Tue|Wed|Thu|Fri|Sat|Sun";
		protected final static String US_DAYS_EEEE = "Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday";
		protected final static String DE_DAYS_E = "Mo|Di|Mi|Do|Fr|Sa|So";
		protected final static String DE_DAYS_EEEE = "Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag";
		protected final static String FR_DAYS_E = "lu|ma|me|je|ve|sa|di";
		protected final static String FR_DAYS_EEEE = "lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche";
		protected final static String US_MONTHS_MMM = "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec";
		protected final static String US_MONTHS_MMMM = "January|February|March|April|May|June|July|August|September|October|November|December";
		protected final static String DE_MONTHS_MMM = "Jan|Feb|Mrz|Apr|Mai|Jun|Jul|Aug|Sep|Okt|Nov|Dez";
		protected final static String DE_MONTHS_MMMM = "Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember";
		protected final static String FR_MONTHS_MMM = "jan|fév|mar|avr|mai|juin|juil|août|sep|oct|nov|déc";
		protected final static String FR_MONTHS_MMMM = "janvier|février|mars|avril|mai|juin|juillet|août|septembre|octobre|novembre|décembre";
		protected final static String US_ZONE_Z = "CET|CEST"; // Central
																// European Time
																// / Summer Time
		protected final static String US_ZONE_ZZZZ = "Central European Time|Central European Summer Time";
		protected final static String DE_ZONE_Z = "MEZ|MESZ"; // Mitteleuropäische
																// Zeit /
																// Sommerzeit
		protected final static String FR_ZONE_Z = "HEC|HESC"; // Mitteleuropäische
		// Zeit /
		// Sommerzeit
		protected final static String DE_ZONE_ZZZZ = "Mitteleuropäische Zeit|Mitteleuropäische Sommerzeit";
		protected final static String US_ERA_G = "BC|AD";
		protected final static String DE_ERA_G = "v. Chr.|n. Chr.";
		protected final static String AM_PM = "AM|PM";

		// Namenskonvention
		protected final static String N_FORMAT = "F_";
		protected final static String N_PATTERN = "P_";
		protected final static String N_PATTERNFULL = "PF_";
		protected final static String N_DATE = "D_";
		protected final static String N_TIME = "T_";
		protected final static String N_DATETIME = "DT_";
		protected final static String N_TIMESTAMP = "TS_";
		protected final static String N_SHORT = "SHORT";
		protected final static String N_LONG = "LONG";

		// Steuerzeichen
		protected final static String P_LINE_BEGIN = "^";
		protected final static String P_LINE_END = "$";

		// Datum
		protected final static String F_DE_D_SHORT = "d.M.y";
		protected final static String P_DE_D_SHORT = "((([1-9]|[12][\\d]|[3][01])\\.([13578]|10|12))|(([1-9]|[12][\\d]|[3][0])\\.([469]|11))|(([1-9]|[12][\\d])\\.2))\\.([\\d]{2})";
		protected final static String PF_DE_D_SHORT = P_LINE_BEGIN + P_DE_D_SHORT + P_LINE_END;

		protected final static String F_DE_D_YEAR = "yyyy";
		protected final static String P_DE_D_YEAR = "(19[\\d]{2}|[2-9][\\d]{3})";
		protected final static String PF_DE_D_YEAR = P_LINE_BEGIN + P_DE_D_YEAR + P_LINE_END;

		protected final static String F_DE_D_YEAR_MONTH = "MM/yyyy";
		protected final static String P_DE_D_YEAR_MONTH = "(0[13578]|10|12)([\\/])(19[\\d]{2}|[2-9][\\d]{3})";
		protected final static String PF_DE_D_YEAR_MONTH = P_LINE_BEGIN + P_DE_D_YEAR_MONTH + P_LINE_END;

		protected final static String F_DE_D_LONG = "dd.MM.yyyy";
		protected final static String P_DE_D_LONG = "(((0[1-9]|[12][\\d]|[3][01])\\.(0[13578]|10|12))|((0[1-9]|[12][\\d]|[3][0])\\.(0[469]|11))|((0[1-9]|[12][\\d])\\.02))\\.(19[\\d]{2}|[2-9][\\d]{3})";
		protected final static String PF_DE_D_LONG = P_LINE_BEGIN + P_DE_D_LONG	+ P_LINE_END;

		protected final static String F_ISO_D_LONG = "yyyy-MM-dd";
		protected final static String P_ISO_D_LONG = "(19[\\d]{2}|[2-9][\\d]{3})-((0[13578]|10|12)-((0[1-9]|[12][\\d]|[3][01]))|((0[469]|11)-(0[1-9]|[12][\\d]|[3][0]))|(02-(0[1-9]|[12][\\d])))";
		protected final static String PF_ISO_D_LONG = P_LINE_BEGIN + P_ISO_D_LONG + P_LINE_END;

		protected final static String F_US_D_SHORT = "M/d/yy";
		protected final static String P_US_D_SHORT = "((([13578]|10|12)\\/([1-9]|[12][\\d]|[3][01]))|(([469]|11)\\/([1-9]|[12][\\d]|[3][0]))|(2\\/([1-9]|[12][\\d])))\\/([\\d]{2})";
		protected final static String PF_US_D_SHORT = P_LINE_BEGIN + P_US_D_SHORT + P_LINE_END;
		
		protected final static String F_US_D_LONG = "MM/dd/yyyy";
		protected final static String P_US_D_LONG = "(((0[13578]|10|12)\\/(0[1-9]|[12][\\d]|[3][01]))|((0[469]|11)\\/(0[1-9]|[12][\\d]|[3][0]))|(02\\/(0[1-9]|[12][\\d])))\\/(19[\\d]{2}|[2-9][\\d]{3})";
		protected final static String PF_US_D_LONG = P_LINE_BEGIN + P_US_D_LONG + P_LINE_END;

		// Zeit
		protected final static String F_DE_T_LONG = "HH:mm:ss";
		protected final static String P_DE_T_LONG = "([01][\\d]|[2][0-4]):([0-5][0-9]|[6][0]):([0-5][0-9]|[6][0])";
		protected final static String PF_DE_T_LONG = P_LINE_BEGIN + P_DE_T_LONG	+ P_LINE_END;

		protected final static String F_ISO_T_LONG = "THH:mm:ss";
		protected final static String P_ISO_T_LONG = "[T]" + P_DE_T_LONG;
		protected final static String PF_ISO_T_LONG = P_LINE_BEGIN + P_ISO_T_LONG + P_LINE_END;

		// Datum + Zeit
		protected final static String F_DE_DT_SHORT = F_DE_D_SHORT + " " + F_DE_T_LONG;
		protected final static String P_DE_DT_SHORT = P_DE_D_SHORT + " " + P_DE_T_LONG;
		protected final static String PF_DE_DT_SHORT = P_LINE_BEGIN	+ P_DE_DT_SHORT + P_LINE_END;
		protected final static String F_DE_DT_LONG = F_DE_D_LONG + " " + F_DE_T_LONG;
		protected final static String P_DE_DT_LONG = P_DE_D_LONG + " " + P_DE_T_LONG;
		protected final static String PF_DE_DT_LONG = P_LINE_BEGIN + P_DE_DT_LONG + P_LINE_END;

		protected final static String F_ISO_DT_LONG = F_ISO_D_LONG + F_ISO_T_LONG;
		protected final static String P_ISO_DT_LONG = P_ISO_D_LONG + P_ISO_T_LONG;
		protected final static String PF_ISO_DT_LONG = P_LINE_BEGIN	+ P_ISO_DT_LONG + P_LINE_END;

		// Zeitstempel
		protected final static String F_DE_TS_SHORT = F_DE_D_SHORT + " " + F_DE_T_LONG + ".SSS";
		protected final static String P_DE_TS_SHORT = P_DE_D_SHORT + " " + P_DE_T_LONG + "\\.[0-9]{3}";
		protected final static String PF_DE_TS_SHORT = P_LINE_BEGIN + P_DE_TS_SHORT + P_LINE_END;
		protected final static String F_DE_TS_LONG = F_DE_D_LONG + " " + F_DE_T_LONG + ".SSSSSS";
		protected final static String P_DE_TS_LONG = P_DE_D_LONG + " " + P_DE_T_LONG + "\\.[0-9]{6}";
		protected final static String PF_DE_TS_LONG = P_LINE_BEGIN + P_DE_TS_LONG + P_LINE_END;

		private static final Map<String, Pattern> compiledPatternMap = new java.util.HashMap<String, Pattern>();
		
		/**
		 * Default-Konstruktor, der nicht au&szlig;erhalb dieser Klasse
		 * aufgerufen werden kann.
		 */
		private DateUtil() {}

		private static Pattern getCompiledPattern(String patternStr) {
			Pattern p = compiledPatternMap.get(patternStr);
			if (p == null) {
				synchronized (compiledPatternMap) {
					Pattern temp = Pattern.compile(patternStr);
					compiledPatternMap.put(patternStr, temp);
					p = temp;
				}
			}
			return p;
		}
		
		/**
		 * Die Methode ermittelt welchem Datumsformat das &uuml;bergebene Datum
		 * entspricht, hierbei werden per Reflection API s&auml;mtliche PF_
		 * Konstanten f&uuml;r die Pr&uuml;fung herangezogen. Trifft ein
		 * Datumsform zu, wird der Namensstamm der Konstante extrahiert und
		 * zur&uuml;ckgegeben.
		 * 
		 * @param date
		 *            Das Datum zu dem der Datumsschl&uuml;ssel der
		 *            Konstantbezeichner ermittelt werden soll.
		 * @return Gibt den Namensstamm der Konstante des zutreffenden
		 *         Datumsformat zur&uuml;ck.
		 */
		private String getDateKey(String date) {
			String dateKey = null;

			Field[] member = this.getClass().getDeclaredFields();

			boolean stop = false;
			List<Field> list = Arrays.asList(member);
			Iterator<Field> iterator = list.iterator();
			while (iterator.hasNext() && !stop) {
				Field field = (Field) iterator.next();
				String fieldName = field.getName();
				if (fieldName.startsWith(N_PATTERNFULL)) {
					try {
						String pattern = (String) field.get(null);
						
						Pattern p = getCompiledPattern(pattern);
						Matcher m = p.matcher(date);
						if (m.matches()) {
							stop = true;

							int position = (fieldName.indexOf("_") + 1);
							dateKey = fieldName.substring(position);
						}
					} catch (IllegalArgumentException e) {
						e.printStackTrace();
					} catch (IllegalAccessException e) {
						e.printStackTrace();
					} catch (SecurityException e) {
						e.printStackTrace();
					}
				}
			}

			return dateKey;
		}

		/**
		 * Pr&uuml;ft ob es sich bei dem &uuml;bergebenen Datum um ein f&uuml;r
		 * diese Klasse parse-bares Datum handelt.
		 * 
		 * @param date
		 *            Das zu &uuml;berpr&uuml;fende Datum.
		 * @return true, wenn das Datum von dieser Klasse unterst&uuml;tzt wird,
		 *         sonst false
		 */
		public boolean isDate(String date) {
			return getDateKey(date) != null ? true : false;
		}

		/**
		 * Die Methode arbeitet mit Reflections und ermittelt den Wert der
		 * Konstanten, die den Namensstamm des &uuml;bergebenen Datums und die
		 * Namenskonvention (Format, Pattern, oder PatternFull) beinhaltet.
		 * 
		 * @param date
		 *            Das Datum zudem der Konstantenwert ermittelt werden soll.
		 * @param norm
		 *            Die Namenskonvention die definiert welcher Wert zum
		 *            jeweiligen Datumsformat zur&uuml;ckgegeben werden soll.
		 * @return Gibt den Wert der Konstanten zur&uuml;ck, die aus den beiden
		 *         Parametern ermittelt wird.
		 */
		private String getFieldValue(String date, String norm) {
			String dateKey = getDateKey(date);
			String fieldKey = null;
			String fieldValue = null;

			if (dateKey != null) {
				fieldKey = norm + dateKey;
				try {
					Field field;
					field = this.getClass().getDeclaredField(fieldKey);
					fieldValue = field.get(null).toString();
				} catch (SecurityException e) {
					e.printStackTrace();
				} catch (NoSuchFieldException e) {
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					e.printStackTrace();
				}
			}

			return fieldValue;
		}

		/**
		 * Ermittelt den regul&auml;ren Ausdruck zu dem dieses Datum passt.
		 * 
		 * @param date
		 *            Das Datum zudem der regul&auml;re Ausdruck ermittelt
		 *            werden soll.
		 * @return Der ermittelte regul&auml;re Ausdruck zum &uuml;bergebenen
		 *         Parameter.
		 */
		@SuppressWarnings("unused")
		private String getPattern(String date) {
			return getFieldValue(date, N_PATTERNFULL);
		}

		/**
		 * Ermittelt das Datumsformat (SimpleDateFormat) f&uuml;r das
		 * &uuml;bergeben Datum.
		 * 
		 * @param date
		 *            Das Datum zudem das Datumsformat ermittelt werden soll.
		 * @return Das ermittelte Datumsformat.
		 */
		private String getFormat(String date) {
			return getFieldValue(date, N_FORMAT);
		}

		/**
		 * Die Methode extrahiert anhand der konfigurierten Datumsformate aus
		 * dem &uuml;bergebenen Text ein Datum, falls das Format
		 * unterst&uuml;tzt wird.
		 * 
		 * @param text
		 *            Das Datum in Textform aus dem ein Java-Date Objekt erzeugt
		 *            werden soll.
		 * @return Gibt ein Java-Date zur&uuml;ck.
		 * @throws ParseException
		 *             Tritt auf wenn das Datum nicht als solches erkannt wird,
		 *             oder es sich nicht um ein Datum handelt.
		 */
		public Date parseDate(String text) throws ParseException {
			Date date = null;
			String format = getFormat(text);

			if (format != null) {
				SimpleDateFormat formater = new SimpleDateFormat(
						format,
						Locale.GERMAN);
				date = formater.parse(text);
			} else {
				StringBuffer supported = new StringBuffer();
				List<String> supportedFormats = getSupportedFormats();
				supported.append("Supported formats are:\n");
				for (String temp : supportedFormats) {
					supported.append(temp);
					supported.append("\n");
				}

				throw new ParseException(text
						+ " is an unsupported date format!\n"
						+ supported.toString(), 0);
			}

			return date;
		}

		/**
		 * @return Gibt eine Liste mit allen aktuell konfigurierten
		 *         Datumsformaten zur&uuml;ck, aus denen automatisch ein Datum
		 *         extrahiert werden kann.
		 */
		protected List<String> getSupportedFormats() {
			List<String> supported = new ArrayList<String>();

			String regex = "^F_[A-Z]{2,3}_(D|DT|TS)_.*";
			Pattern pattern = getCompiledPattern(regex);

			Field[] member = this.getClass().getDeclaredFields();
			List<Field> list = Arrays.asList(member);
			Iterator<Field> iterator = list.iterator();
			while (iterator.hasNext()) {
				Field field = (Field) iterator.next();
				String fieldName = field.getName();

				Matcher matcher = pattern.matcher(fieldName);
				if (matcher.matches()) {
					try {
						String value;
						value = (String) field.get(null);
						supported.add(value);
					} catch (IllegalArgumentException e) {
						e.printStackTrace();
					} catch (IllegalAccessException e) {
						e.printStackTrace();
					}
				}
			}

			return supported;
		}

		/**
		 * @return Liefert die einzige Instanz dieser Klasse.
		 */
		public static DateUtil getInstance() {
			return instance;
		}
	}

	/**
	 * Die Methode ermittelt ob das Datum dem vorgegebenen Datumsformat
	 * entspricht und somit in dieses geparsed werden kann.
	 * 
	 * @param pattern
	 *            Das vorgegebene Datumsformat als regul&auml;rer Ausdruck.
	 * @param date
	 *            Das Datum als String das zu validieren ist.
	 * @return true, wenn das Datum dem Format entspricht, sonst false
	 * 
	 *         {Category} GenericDateUtils {talendTypes} boolean {param}
	 *         string("") pattern: String {param} string("") date: String
	 *         {example} matches("dd.MM.yyyy", "21.05.2013") # true {example}
	 *         matches("dd/mm/yyyy", "21.05.2013") # false
	 */
	public static boolean matches(String pattern, String date) {
		Pattern p = DateUtil.getCompiledPattern(pattern);
		Matcher m = p.matcher(date);

		return m.matches();
	}

	/**
	 * Pr&uuml;ft ob es sich bei dem &uuml;bergebenen Datum um ein f&uuml;r
	 * diese Klasse parse-bares Datum handelt.
	 * 
	 * @param date
	 *            Das zu &uuml;berpr&uuml;fende Datum.
	 * @return true, wenn das Datum von dieser Klasse unterst&uuml;tzt wird,
	 *         sonst false
	 * 
	 *         {Category} GenericDateUtils {talendTypes} boolean {param}
	 *         string("21.05.2013") date: String {example} isDate("21.05.2013")
	 *         # true {example} isDate("Hello world!") # false
	 */
	public static boolean isDate(String date) {
		DateUtil instance = DateUtil.getInstance();
		return instance.isDate(date);
	}

	/**
	 * Die Methode extrahiert anhand der konfigurierten Datumsformate aus dem
	 * &uuml;bergebenen Text ein Datum, falls das Format unterst&uuml;tzt wird.
	 * 
	 * @param text
	 *            Das Datum in Textform aus dem ein Java-Date Objekt erzeugt
	 *            werden soll.
	 * @return Gibt ein Java-Date zur&uuml;ck.
	 * @throws ParseException
	 *             Tritt auf wenn das Datum nicht als solches erkannt wird, oder
	 *             es sich nicht um ein Datum handelt.
	 * 
	 *             {Category} GenericDateUtils {talendTypes} Date {param}
	 *             string("21.05.2013") text: String {example}
	 *             parseDate("21.05.2013") # Date {example}
	 *             parseDate("05/21/2013") # Date {example}
	 *             parseDate("21052013T134251") # ParseException
	 */
	public static Date parseDate(String text) throws ParseException {
		DateUtil instance = DateUtil.getInstance();
		return instance.parseDate(text);
	}

	/**
	 * @return Gibt eine Liste mit allen aktuell konfigurierten Datumsformaten
	 *         zur&uuml;ck, aus denen automatisch ein Datum extrahiert werden
	 *         kann.
	 * 
	 *         {Category} GenericDateUtils {talendTypes} Object {example}
	 *         getSupportedFormats()
	 */
	public static List<String> getSupportedFormats() {
		DateUtil instance = DateUtil.getInstance();
		return instance.getSupportedFormats();
	}
	
}