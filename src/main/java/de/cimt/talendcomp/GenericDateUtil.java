package de.cimt.talendcomp;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

/**
 * Utility class to parse a String into a Date 
 * by testing a number of common pattern
 * This class is thread save.
 * 
 * @author jan.lolling@cimt-ag.de
 */
public class GenericDateUtil {
	
	private static ThreadLocal<DateParser> threadLocal = new ThreadLocal<DateParser>();
	
    /**
     * parseDuration: returns the Date from the given text representation containing the time part as duration
     * Uses an internal list of patterns to parse the source.
     * 
     * @param source the formatted time as String
     * @return Date object representing the duration
     */
	public static Date parseDuration(String source) throws ParseException {
		return parseDuration(source, (String[]) null);
	}

    /**
     * parseDuration: returns the Date from the given source as duration
     * Tolerates if the content does not fit to the given pattern and retries it
     * with build in patterns
     * @param source the formatted time as String
     * @param suggestedPattern an array of suggested patterns
     * @return Date object representing the duration
     */
	public static Date parseDuration(String source, String ...suggestedPattern) throws ParseException {
		return getDateParser().parseDuration(source, suggestedPattern);
	}

    /**
     * parseDuration: returns the Date from the given source as duration
     * Tolerates if the content does not fit to the given pattern and retries it
     * with build in patterns
     * @param source date or time as Double in which a 1 is one day
     * @return Long object representing the duration
     */
	public static Long parseDuration(Double source) {
		return getDateParser().getDuration(source);
	}

    /**
     * parseDate: returns the Date from the given text representation
     * Uses an internal list of patterns to parse the source.
     * 
     * @param source the formatted date as String
     * @return Date object representing the Date
     */
	public static Date parseDate(String source) throws ParseException {
		return parseDate(source, (String[]) null);
	}

	/**
     * parseDate: returns the Date from the given text representation
     * Tolerates if the content does not fit to the given pattern and retries it
     * with build in patterns
     * 
     * @param source the formatted time as String
     * @param suggestedPattern an array of suggested patterns
     * @return Date object representing the Date
     */
	public static Date parseDate(String source, String ...suggestedPattern) throws ParseException {
		return getDateParser().parseDate(source, suggestedPattern);
	}
	
	private static DateParser getDateParser() {
		DateParser p = threadLocal.get();
		if (p == null) {
			p = new DateParser();
			threadLocal.set(p);
		}
		return p;
	}
	
	private static class DateParser {
		
		private List<String> datePatternList = null;
		private List<String> timePatternList = null;
		
		private static final int SECONDS_PER_MINUTE = 60;
		private static final int MINUTES_PER_HOUR = 60;
		private static final int HOURS_PER_DAY = 24;
		private static final int SECONDS_PER_DAY = (HOURS_PER_DAY * MINUTES_PER_HOUR * SECONDS_PER_MINUTE);
		private static final long DAY_MILLISECONDS = SECONDS_PER_DAY * 1000L;
		
		DateParser() {
			datePatternList = new ArrayList<String>();
			datePatternList.add("yyyy-MM-dd");
			datePatternList.add("dd.MM.yyyy");
			datePatternList.add("d.MM.yyyy");
			datePatternList.add("d.M.yy");
			datePatternList.add("dd.MM.yy");
			datePatternList.add("dd.MMM.yyyy");
			datePatternList.add("MM/dd/yyyy");
			datePatternList.add("MM/dd/yy");
			datePatternList.add("M/d/yy");
			datePatternList.add("dd/MM/yyyy");
			datePatternList.add("dd/MM/yy");
			datePatternList.add("dd/MMM/yyyy");
			datePatternList.add("dd. MMMM yyyy");
			datePatternList.add("dd. MMM yyyy");
			datePatternList.add("MMMM dd'th' yyyy");
			datePatternList.add("MMM dd'th' yyyy");
			datePatternList.add("dd'th' MMMM yyyy");
			datePatternList.add("dd'th' MMM yyyy");
			datePatternList.add("dd-MM-yyyy");
			datePatternList.add("dd-MM-yy");
			datePatternList.add("dd-MMM-yyyy");
			datePatternList.add("d-M-yy");
			datePatternList.add("yyyyMMdd");
			datePatternList.add("yyyyMM");
			datePatternList.add("yyyy");
			timePatternList = new ArrayList<String>();
			timePatternList.add(" mm'′'ss'″'");
			timePatternList.add(" mm''ss'\"'");
			timePatternList.add(" HH'h'mm'm'ss's'");
			timePatternList.add(" HH'h'mm'm'");
			timePatternList.add(" mm'm'ss's'");
			timePatternList.add("'T'HH:mm:ss.SSSZ");
			timePatternList.add("'T'HH:mm:ss.SSS");
			timePatternList.add(" HH:mm:ss.SSS");
			timePatternList.add(" HH:mm:ss");
			timePatternList.add(" mm:ss");
			timePatternList.add(" HHmmss");
			timePatternList.add(" mmss");
		}
		
		private Date parseDate(String text, String ... userPattern) throws ParseException {
			if (text != null && text.trim().isEmpty() == false) {
				Date dateValue = null;
				if (userPattern != null) {
					for (int i = userPattern.length - 1; i >= 0; i--) {
						if (datePatternList.contains(userPattern[i])) {
							datePatternList.remove(userPattern[i]);
						}
						datePatternList.add(0, userPattern[i]);
					}
				}
				SimpleDateFormat sdf = new SimpleDateFormat();
				for (String pattern : datePatternList) {
					if (pattern != null) {
						sdf.applyPattern(pattern.trim());
						try {
							dateValue = sdf.parse(text);
							// if we continue here the pattern fits
							// now we know the date is correct, lets try the time part:
							if (text.length() - pattern.length() >= 6) {
								// there is more in the text than only the date
								for (String timepattern : timePatternList) {
									String dateTimePattern = pattern + timepattern;
									sdf.applyPattern(dateTimePattern);
									try {
										dateValue = sdf.parse(text);
										// we got it
										pattern = dateTimePattern;
										break;
									} catch (ParseException e1) {
										// ignore parsing errors, we are trying
									}
								}
							}
							return dateValue;
						} catch (ParseException e) {
							// the pattern obviously does not work
							continue;
						}
					}
				}
				throw new ParseException("The value: " + text + " could not be parsed to a Date.", 0);
			} else {
				return null;
			}
		}

		private Date parseDuration(String text, String ... userPattern) throws ParseException {
			if (text != null && text.trim().isEmpty() == false) {
				Date timeValue = null;
				if (userPattern != null) {
					for (int i = userPattern.length - 1; i >= 0; i--) {
						if (timePatternList.contains(userPattern[i])) {
							timePatternList.remove(userPattern[i]);
						}
						timePatternList.add(0, userPattern[i]);
					}
				}
				SimpleDateFormat sdf = new SimpleDateFormat();
				sdf.setTimeZone(getUTCTimeZone());
				for (String pattern : timePatternList) {
					if (pattern != null) {
						sdf.applyPattern(pattern.trim());
						try {
							timeValue = sdf.parse(text);
							// if we continue here the pattern fits
							return timeValue;
						} catch (ParseException e) {
							// the pattern obviously does not work
							continue;
						}
					}
				}
				throw new ParseException("The value: " + text + " could not be parsed to a Date as duration.", 0);
			} else {
				return null;
			}
		}
		
		private Long getDuration(Double timeInExcel) {
			if (timeInExcel != null) {
		        int wholeDays = (int) Math.floor(timeInExcel);
		        int millisecondsInDay = (int) ((timeInExcel - wholeDays) * DAY_MILLISECONDS + 0.5);
		        Calendar cal = Calendar.getInstance(getUTCTimeZone());
		        cal.setTimeInMillis(0);
		        cal.set(Calendar.DAY_OF_YEAR, wholeDays + 1);
		        cal.set(Calendar.MILLISECOND, millisecondsInDay);
		        return cal.getTimeInMillis();
			} else {
				return null;
			}
		}

	}
	
    private static java.util.TimeZone utcTimeZone = null;

    private static java.util.TimeZone getUTCTimeZone() {
    	if (utcTimeZone == null) {
    		utcTimeZone = java.util.TimeZone.getTimeZone("UTC");
    	}
    	return utcTimeZone;
    }

}