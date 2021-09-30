package com.dfl.report.common;

import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCPreferenceService.TCPreferenceLocation;
import com.teamcenter.rac.kernel.TCPreferenceService.TCPreferenceProtectionScope;
import com.teamcenter.rac.kernel.TCPreferenceService.TCPreferenceType;

public final class TCPreferenceServiceUtils {
	public static String[] getPrefernceValuesNoTips(String preferenceName, String[] defaultValue) throws Exception {
		TCPreferenceService tcps = TCUtils.getPreferenceService();

		String[] value = tcps.getStringArray(0, preferenceName);
		if ((value != null) && (value.length > 0)) {
			return value;
		}
		return defaultValue;
	}

	public static String[] getPrefernceValues(String preferenceName, String[] defaultValue) throws Exception {
		TCPreferenceService tcps = TCUtils.getPreferenceService();

		String[] value = tcps.getStringArray(0, preferenceName);

		if ((value != null) && (value.length > 0)) {
			return value;
		}
		EclipseUtils.info(Messages.preference + "[" + preferenceName + "]" + Messages.unexist);

		return defaultValue;
	}

	public static String getPrefernceValue(String preferenceName, String defaultValue) throws Exception {
		TCPreferenceService tcps = TCUtils.getPreferenceService();

		String value = tcps.getString(0, preferenceName);

		if ((value != null) && (value.length() > 0)) {
			return value;
		}
		EclipseUtils.info(Messages.preference + "[" + preferenceName + "]" + Messages.unexist);

		return defaultValue;
	}

	public static String getStringPreferenceValue(String preferenceName) throws Exception {
		TCPreferenceService tcps = TCUtils.getPreferenceService();
		return tcps.getStringValue(preferenceName);
	}

	public static String[] getStringPreferenceValues(String preferenceName) throws Exception {
		TCPreferenceService tcps = TCUtils.getPreferenceService();
		return tcps.getStringValues(preferenceName);
	}

	public static int getIntPreferenceValue(String preferenceName) throws Exception {
		TCPreferenceService tcps = TCUtils.getPreferenceService();
		return tcps.getIntegerValue(preferenceName).intValue();
	}

	public static int getUserIntPreferenceValue(String preferenceName) throws Exception {
		TCPreferenceService tcps = TCUtils.getPreferenceService();
		return tcps.getIntegerValueAtLocation(preferenceName, TCPreferenceService.TCPreferenceLocation.USER_LOCATION)
				.intValue();
	}

//	public static void createUserIntPreferenceValue(String preferenceName, String preferenceDesc,
//			String preferenceCategory, boolean isArray, boolean isEnvEnabled, int preferenceValue) throws Exception {
//		TCPreferenceService tcps = TCUtils.getPreferenceService();
//		TCPreferenceService.TCPreferenceProtectionScope proptectionScope = TCPreferenceService.TCPreferenceProtectionScope.USER_PROTECTION_SCOPE;
//		TCPreferenceService.TCPreferenceLocation prefLocation = TCPreferenceService.TCPreferenceLocation.USER_LOCATION;
//		TCPreferenceService.TCPreferenceType prefType = TCPreferenceService.TCPreferenceType.INTEGER_TYPE;
//		tcps.create(preferenceName, proptectionScope, preferenceDesc, preferenceCategory, prefType, isArray,
//				isEnvEnabled, prefLocation, preferenceValue, new String[] { preferenceValue });
//	}

	public static void setUserIntPreferenceValue(String preferenceName, int preferenceValue) throws Exception {
		TCPreferenceService tcps = TCUtils.getPreferenceService();
		tcps.setIntegerValue(preferenceName, Integer.valueOf(preferenceValue));
	}
}
