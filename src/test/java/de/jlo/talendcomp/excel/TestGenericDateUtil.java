package de.jlo.talendcomp.excel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.junit.Test;

import de.jlo.talendcomp.excel.GenericDateUtil.DateParser;

public class TestGenericDateUtil {
	
	@Test
	public void testTime() throws ParseException {
		String s = "4m 55s";
		Date result = GenericDateUtil.parseDuration(s, "mm'm'ss's'");
		long actual = result.getTime();
		long expected = 295000l;
		System.out.println("(1) Time in ms: " + actual);
		assertEquals(expected, actual);
		s = "4'55\""; // 4'55"
		result = GenericDateUtil.parseDuration(s, (String) null); 
		actual = result.getTime();
		System.out.println("(2) Time in ms: " + actual);
		assertEquals(expected, actual);
		s = "4' 55\""; // 4'55"
		result = GenericDateUtil.parseDuration(s, "HH:mm:ss"); 
		actual = result.getTime();
		System.out.println("(3) Time in ms: " + actual);
		assertEquals(expected, actual);
		s = "4' 55”"; // 4'55"
		result = GenericDateUtil.parseDuration(s, "HH:mm:ss"); 
		actual = result.getTime();
		System.out.println("(3) Time in ms: " + actual);
		assertEquals(expected, actual);
		s = "4' 55“"; // 4'55"
		result = GenericDateUtil.parseDuration(s, "HH:mm:ss"); 
		actual = result.getTime();
		System.out.println("(3) Time in ms: " + actual);
		assertEquals(expected, actual);
		s = "00:04:55"; 
		result = GenericDateUtil.parseDuration(s, "HH:mm:ss");
		actual = result.getTime();
		System.out.println("(4) Time in ms: " + actual);
		assertEquals(expected, actual);
		s = "23:59"; 
		result = GenericDateUtil.parseDuration(s, "HH:mm:ss");
		expected = 1439000l;
		actual = result.getTime();
		System.out.println("(5) Time in ms: " + actual);
		assertEquals(expected, actual);
		s = "4m 55s";
		result = GenericDateUtil.parseDuration(s, "HH'h'mm'm'ss's'");
		actual = result.getTime();
		expected = 295000l;
		System.out.println("(6) Time in ms: " + actual);
		assertEquals(expected, actual);
		s = "4′ 55″";
		result = GenericDateUtil.parseDuration(s, (String) null); 
		actual = result.getTime();
		System.out.println("(7) Time in ms: " + actual);
		assertEquals(expected, actual);
		s = "0455";
		result = GenericDateUtil.parseDuration(s, (String) null); 
		actual = result.getTime();
		System.out.println("(8) Time in ms: " + actual);
		assertEquals(expected, actual);
		s = "000455";
		result = GenericDateUtil.parseDuration(s, (String) null); 
		actual = result.getTime();
		System.out.println("(9) Time in ms: " + actual);
		assertEquals(expected, actual);
		s = "13:00:00";
		result = GenericDateUtil.parseDuration(s, (String) null); 
		actual = result.getTime();
		System.out.println("(9) Time in ms: " + actual);
		assertEquals(46800000l, actual);
		s = "01:00:00";
		result = GenericDateUtil.parseDuration(s, (String) null); 
		actual = result.getTime();
		System.out.println("(9) Time in ms: " + actual);
		assertEquals(3600000l, actual);
	}
	
	@Test
	public void testDateAndTime() throws ParseException {
		String s = "2016-12-11 13:26:11";
		Long actual = GenericDateUtil.parseDate(s, "yyyy-MM-dd HH:mm:ss").getTime();
		Long expected = 1481459171000l;
		assertEquals(expected, actual);
	}
	
	
	@Test
	public void testDate() throws Exception {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		String s2 = "01.03.2017";
		Date date2 = GenericDateUtil.getDateParser(false).parseDate(s2, (String) null);
		String s1 = "01.03.2017";
		Date date1 = GenericDateUtil.getDateParser(false).parseDate(s1, "dd.MM.yy");
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "2017-03-01";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "03/01/2017";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, Locale.ENGLISH, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "03/01/17";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, Locale.ENGLISH, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "01.03.17";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, Locale.ENGLISH, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "1th Mar 2017";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, Locale.ENGLISH, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "Mar 1th 2017";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "01. March 2017";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "March 2017";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "03/2017";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, "MM/yyyy");
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "01. März 2017";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, new Locale("de"), (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "März 2017";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, Locale.GERMANY, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "KW 9/2017";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, Locale.GERMANY, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.before(date2));
		s1 = "w/c 9.2017";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, Locale.ENGLISH, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.before(date2));
		s1 = "01.03.17";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, Locale.ENGLISH, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "01-03-2017";
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, Locale.ENGLISH, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "2017";
		s2 = "01.01.2017";
		date2 = GenericDateUtil.getDateParser(false).parseDate(s2, (String) null);
		date1 = GenericDateUtil.getDateParser(false).parseDate(s1, Locale.ENGLISH, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
	}

	@Test
	public void testZeroDate() throws ParseException {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String s1 = "0000-00-00 05:00:00";
//		Date date1 = GenericDateUtil.parseDate(s1, Locale.ENGLISH, (String) null);#
		Date date1 = sdf.parse(s1);
		System.out.println("date1: " + date1);
		System.out.println("date1 ms: " + date1.getTime());
		assertTrue(GenericDateUtil.isZeroDate(date1));
	}

	@Test
	public void testInvalidDate() throws ParseException {
		String s = "2016-13-11 13:26:11";
		DateParser p =  GenericDateUtil.getDateParser(false);
		try {
			p.setLenient(false);
			Date actual = p.parseDate(s, "yyyy-MM-dd HH:mm:ss");
			System.out.println(actual);
			assertTrue(false);
		} catch (Exception e) {
			e.printStackTrace();
			assertTrue(true);
		}
	}

}
