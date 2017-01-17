package de.cimt.talendcomp;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
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
	
	public static Date parseDate(String source) throws ParseException {
		return parseDate(source, (String[]) null);
	}

    /**
     * parseDate: returns the Date from the given text representation
     * Tolerates if the content does not fit to the given pattern and retries it
     * with build in patterns
     * 
     * {Category} GenericDateUtil
     * 
     * {talendTypes} Date
     * 
     * {param} String(dateString)
     * {param} String(suggestedPattern)
     * 
     * {example} parseDate(dateString, suggestedPattern).
     */
	public static Date parseDate(String source, String ...suggestedPattern) throws ParseException {
		DateParser p = threadLocal.get();
		if (p == null) {
			p = new DateParser();
			threadLocal.set(p);
		}
		return p.parseDate(source, suggestedPattern);
	}
	
    /**
     * parseTime: returns the Date from the given text representation which consists only the time part
     * Tolerates if the content does not fit to the given pattern and retries it
     * with build in patterns
     * 
     * {Category} GenericDateUtil
     * 
     * {talendTypes} Date
     * 
     * {param} String(timeString)
     * {param} String(suggestedPattern)
     * 
     * {example} parseTime(timeString, suggestedPattern).
     */
	public static Date parseTime(String timeString, String ...suggestedPattern) throws ParseException {
		DateParser p = threadLocal.get();
		if (p == null) {
			p = new DateParser();
			threadLocal.set(p);
		}
		return p.parseTime(timeString, suggestedPattern);
	}

	static class DateParser {
		
		private List<String> datePatternList = null;
		private List<String> timePatternList = null;
		
		DateParser() {
			datePatternList = new ArrayList<String>();
			datePatternList.add("yyyy-MM-dd");
			datePatternList.add("dd.MM.yyyy");
			datePatternList.add("d.MM.yyyy");
			datePatternList.add("d.M.yy");
			datePatternList.add("dd.MM.yy");
			datePatternList.add("dd.MMM.yyyy");
			datePatternList.add("yyyyMMdd");
			datePatternList.add("dd/MM/yyyy");
			datePatternList.add("dd/MM/yy");
			datePatternList.add("dd/MMM/yyyy");
			datePatternList.add("d/M/yy");
			datePatternList.add("MM/dd/yyyy");
			datePatternList.add("MM/dd/yy");
			datePatternList.add("dd/MMM/yyyy");
			datePatternList.add("M/d/yy");
			datePatternList.add("dd-MM-yyyy");
			datePatternList.add("dd-MM-yy");
			datePatternList.add("dd-MMM-yyyy");
			datePatternList.add("d-M-yy");
			datePatternList.add("yyyyMM");
			datePatternList.add("yyyy");
			timePatternList = new ArrayList<String>();
			timePatternList.add("'T'HH:mm:ss.SSSZ");
			timePatternList.add(" HHmmss");
			timePatternList.add(" HH'h'mm'm'ss's'");
			timePatternList.add(" HH'h' mm'm' ss's'");
			timePatternList.add(" HH:mm:ss.SSS");
			timePatternList.add(" HH:mm:ss");
			timePatternList.add(" mm''ss'\"'");
			timePatternList.add(" HH'h'mm'm'");
			timePatternList.add(" HH'h' mm'm'");
		}
		
		public Date parseDate(String text, String ... userPattern) throws ParseException {
			if (text != null) {
				SimpleDateFormat sdf = new SimpleDateFormat();
				Date dateValue = null;
				if (userPattern != null) {
					for (int i = userPattern.length - 1; i >= 0; i--) {
						if (datePatternList.contains(userPattern[i])) {
							datePatternList.remove(userPattern[i]);
						}
						datePatternList.add(0, userPattern[i]);
					}
				}
				for (String pattern : datePatternList) {
					sdf.applyPattern(pattern);
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
				throw new ParseException("The value: " + text + " could not be parsed to a Date.", 0);
			} else {
				return null;
			}
		}

		public Date parseTime(String text, String ... userPattern) throws ParseException {
			SimpleDateFormat sdf = new SimpleDateFormat();
			sdf.setTimeZone(getUTCTimeZone());
			Date timeValue = null;
			if (userPattern != null) {
				for (int i = userPattern.length - 1; i >= 0; i--) {
					if (timePatternList.contains(userPattern[i])) {
						timePatternList.remove(userPattern[i]);
					}
					timePatternList.add(0, userPattern[i]);
				}
			}
			for (String pattern : timePatternList) {
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
			throw new ParseException("The value: " + text + " could not be parsed to a Date (only with time).", 0);
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