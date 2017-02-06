package de.cimt.talendcomp.excel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.junit.Test;

import de.cimt.talendcomp.GenericDateUtil;

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
	}
	
	@Test
	public void testDate() throws ParseException {
		String s = "2016-12-11 13:26:11";
		Long actual = GenericDateUtil.parseDate(s, "yyyy-MM-dd HH:mm:ss").getTime();
		Long expected = 1481459171000l;
		assertEquals(expected, actual);
	}
	
	@Test
	public void testDateYearAndFull() throws Exception {
		String s1 = "2017";
		String s2 = "01.03.2017";
		Date date1 = GenericDateUtil.parseDate(s1, (String) null);
		Date date2 = GenericDateUtil.parseDate(s2, "dd.MM.yyyy");
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.before(date2));
		s1 = "01.01.2017";
		date1 = GenericDateUtil.parseDate(s1, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.before(date2));
		s1 = "2017-01-01";
		date1 = GenericDateUtil.parseDate(s1, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.before(date2));
		s1 = "01/14/2017";
		date1 = GenericDateUtil.parseDate(s1, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.before(date2));
		s1 = "14th Jan 2017";
		date1 = GenericDateUtil.parseDate(s1, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.before(date2));
		s1 = "Mar 1th 2017";
		date1 = GenericDateUtil.parseDate(s1, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "01. March 2017";
		date1 = GenericDateUtil.parseDate(s1, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "March 2017";
		date1 = GenericDateUtil.parseDate(s1, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "03/2017";
		date1 = GenericDateUtil.parseDate(s1, "MM/yyyy");
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "01. März 2017";
		date1 = GenericDateUtil.parseDate(s1, Locale.GERMANY, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "März 2017";
		date1 = GenericDateUtil.parseDate(s1, Locale.GERMANY, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.equals(date2));
		s1 = "KW 9/2017";
		date1 = GenericDateUtil.parseDate(s1, Locale.GERMANY, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.before(date2));
		s1 = "w/c 9.2017";
		date1 = GenericDateUtil.parseDate(s1, Locale.ENGLISH, (String) null);
		System.out.println("date1: " + sdf.format(date1));
		System.out.println("date2: " + sdf.format(date2));
		assertTrue(date1.before(date2));
	}

}
