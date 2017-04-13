/**
 * Copyright 2015 Jan Lolling jan.lolling@gmail.com
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package de.jlo.talendcomp.excel;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * <p>Title: StringReplacer</p>
 * <p>Description: tauscht in Strings Inhalte aus. Das Austauschen erfolgt über StringBuffer
 * und kommt deshalb ohne String-Kopieraktionen aus.</p>
 * <p>Anwendungsbeispiel:
 * <blockquote><pre>
 *   private String html2Unicode(String text) {
 *       StringReplacer replacer = new StringReplacer(text);
 *       replacer.replace("%3C","<");
 *       replacer.replace("%3E",">");
 *       replacer.replace("+"," ");
 *       replacer.replace("%25","%");
 *       replacer.replace("equals","=");
 *       return replacer.getResultText();
 *   }
 * </pre></blockquote>
 * </p>
 * @author Jan Lolling
 * @version 1.0
 */
public class StringReplacer {

    private String originalText;    // Original
    private StringBuilder sb;        // Puffer in dem das Ergebnis landet
    private int p0;
    private int p1;
    private int searchTextLength;

    /**
     * Constructor
     * @param originalText Text in dem die Erstzungen durchgefährt werden sollen
     */
    public StringReplacer(String originalText) {
        this.originalText = originalText;
        sb = new StringBuilder();
    }

    /**
     * diese Methode ersetzt den Suchtext im Originaltext durch den replaceText
     * @param searchText - zu suchender Text
     * @param replaceText - Ersatz für den zu suchenden Text
     * @return true wenn Austausch erfolgte, false wenn Suchtext nicht gefunden.
     */
    public boolean replace(String searchText, String replaceText) {
    	return replace(searchText, replaceText, true);
    }

    /**
     * diese Methode ersetzt den Suchtext im Originaltext durch den replaceText
     * @param searchText - zu suchender Text
     * @param replaceText - Ersatz für den zu suchenden Text
     * @param caseSensitive true replace text will be found case insensitive 
     * @return true wenn Austausch erfolgte, false wenn Suchtext nicht gefunden.
     */
    public boolean replace(String searchText, String replaceText, boolean caseSensitive) {
        p0 = 0;
        p1 = 0;
        boolean ok = false;
        if ((searchText != null) && (searchText.length() > 0)) {
            searchTextLength = searchText.length();
            sb.setLength(0); // StringBuffer leeren
            while (true) {
                // als erstes den Suchtext finden.
            	if (caseSensitive) {
                    p1 = originalText.indexOf(searchText, p0);
            	} else {
                    p1 = originalText.toLowerCase().indexOf(searchText.toLowerCase(), p0);
            	}
                if (p1 != -1) {
                    // gefunden !
                    // den Originaltext bis zum Anfang des gefunden Suchtextes in den StringBuffer speichern
                    sb.append(originalText.substring(p0, p1));
                    if (replaceText != null) {
                        sb.append(replaceText); // den Ersatztext hinzufügen
                    }
                    p0 = p1 + searchTextLength; // den Start für die nächste Suche festlegen
                    ok = true;
                } else if (ok) {
                    // wenn nichts ersetzbares mehr gefunden wird, dann den Rest noch anhängen
                    // aber alles nur wenn zuvor etwas gefunden wurde.
                    sb.append(originalText.substring(p0, originalText.length()));
                    break;
                } else {
                    break;
                }
            } // while (p1 != -1)
        } // if ((searchText != null) && (searchText.length()>0))
        if (ok) {
            originalText = sb.toString();
        }
        return ok;
    }

    /**
     * bringt den Originaltext mit den Ersetzungen zurück
     * @return bearbeiteter Text
     */
    public String getResultText() {
        return originalText;
    }

	public static String extractByRegexGroups(String content, String regex) {
		if (regex != null) {
			if (content != null) {
				content = content.trim();
				final StringBuffer sb = new StringBuffer();
				Pattern pattern = Pattern.compile(regex);
		        Matcher matcher = pattern.matcher(content);
		        while (matcher.find()) {
		            if (matcher.start() < matcher.end()) {
		            	for (int i = 1, n = matcher.groupCount(); i <= n; i++) {
			                sb.append(matcher.group(i));
		            	}
		            }
		        }
				return sb.toString();
			} else {
				return null;
			}
		} else {
			return content;
		}
	}

}
