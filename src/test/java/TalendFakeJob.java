import java.util.HashMap;
import java.util.Map;

import routines.system.RunStat;

public class TalendFakeJob {

	static Map<String, Object> globalMap = new HashMap<String, Object>();
	static String currentComponent = "";
	
	static final String jobVersion = "0.1";
	static final String jobName = "";
	static final String projectName = "COMPDEV";
	static public Integer errorCode = null;

	static final java.util.Map<String, Long> start_Hash = new java.util.HashMap<String, Long>();
	static final java.util.Map<String, Long> end_Hash = new java.util.HashMap<String, Long>();
	static final java.util.Map<String, Boolean> ok_Hash = new java.util.HashMap<String, Boolean>();
	static public final java.util.List<String[]> globalBuffer = new java.util.ArrayList<String[]>();

	static RunStat runStat = new RunStat();
	
	static boolean execStat = false;
	static String iterateId = "";
	static java.util.Map<String, Object> resourceMap = new java.util.HashMap<String, Object>();
}
