package com.legao.word.poi;

import org.springframework.util.StringUtils;
import java.util.ArrayList;
import java.util.List;

public class WordStringUtil {
    public static String removeStr(String src, String str) {
        if (src == null || str == null) return src;
        int idx = src.indexOf(str);
        if (idx == -1) return src;
        int pst = 0;
        char[] cs = src.toCharArray();
        char[] rs = new char[src.length() - str.length()];
        for (int i = 0; i < cs.length; i++) {
            if (i >= idx && i < idx + str.length()) continue;
            rs[pst] = cs[i];
            pst++;
        }
        return new String(rs);
    }

    public static String replaceStr(String src, String target, String replacement) {
        if (src == null || target == null || replacement == null) return src;
        int idx = src.indexOf(target);
        if (idx == -1) return src;
        int pst = 0;
        char[] cs = src.toCharArray();
        char[] rs = new char[src.length() - target.length() + replacement.length()];
        for (int i = 0; i < cs.length; i++) {
            if (i == idx) {
                for (char c : replacement.toCharArray()) {
                    rs[pst] = c;
                    pst++;
                }
                continue;
            }
            if (i > idx && i < idx + target.length()) continue;
            rs[pst] = cs[i];
            pst++;
        }
        return new String(rs);
    }

    /**
     * @param src
     * @param target
     * @param replacement
     * @return
     */
    public static String replaceAllStr(String src, String target, String replacement) {
        if (src == null || target == null || replacement == null) return src;
        int idx = src.indexOf(target);
        if (idx == -1) return src;
        int pst = 0;
        char[] cs = src.toCharArray();
        char[] rs = new char[src.length() - target.length() + replacement.length()];
        for (int i = 0; i < cs.length; i++) {
            if (i == idx) {
                for (char c : replacement.toCharArray()) {
                    rs[pst] = c;
                    pst++;
                }
                continue;
            }
            if (i > idx && i < idx + target.length()) continue;
            rs[pst] = cs[i];
            pst++;
        }
        return replaceStr(new String(rs), target, replacement);
    }


    public static List<String> splitString(String src, String target) {

        List<String> result = new ArrayList<>();
        String tar = turnSpecialString(target);
        src = StringUtils.replace(src, target, target + " ");
        String[] s = src.split(tar);
        if (s.length == 0 && src.contains(target)) {
            int number = src.length() / target.length();
            while (number != 0) {
                result.add(target);
                number--;
            }
        }
        if (!src.contains(target)) {
            result.add(src);
            return result;
        }
        for (int i = 0; i < s.length; i++) {
            s[i] = replaceAllStr(s[i], target + " ", target);
            if ("".equals(s[i])) {
                result.add(target);
                continue;
            }
            result.add(s[i]);
            if (i == s.length - 1 && !src.endsWith(target)) {
                continue;
            }
            result.add(target);

        }
        return result;
    }


    public static List<String> splitStrings(List<String> src, String replacement) {
        List<String> results = new ArrayList<>();
        for (String string : src) {
            List<String> sdsd = splitString(string, replacement);
            results.addAll(sdsd);
        }
        return results;
    }


    public static List<String> splitStringList(List<String> src, List<String> replacements) {
        List<String> results = new ArrayList<>();

        for (String replacement : replacements) {
            List<String> sdsd = splitStrings(src, replacement);

            src = sdsd;
            results = sdsd;
        }

        return results;
    }


    public static String turnSpecialString(String str) {
        String resStr = "";
        for (int i = 0; i < str.length(); i++) {
            //转义字符 ([{ / ^ -$ ¦ } ] ) ? *+ .
            if ((String.valueOf(str.charAt(i))).equals("(")
                    || (String.valueOf(str.charAt(i))).equals("[")
                    || (String.valueOf(str.charAt(i))).equals("{")
                    || (String.valueOf(str.charAt(i))).equals("/")
                    || (String.valueOf(str.charAt(i))).equals("^")
                    || (String.valueOf(str.charAt(i))).equals("-")
                    || (String.valueOf(str.charAt(i))).equals("$")
                    || (String.valueOf(str.charAt(i))).equals("¦")
                    || (String.valueOf(str.charAt(i))).equals("}")
                    || (String.valueOf(str.charAt(i))).equals("]")
                    || (String.valueOf(str.charAt(i))).equals(")")
                    || (String.valueOf(str.charAt(i))).equals("?")
                    || (String.valueOf(str.charAt(i))).equals("*")
                    || (String.valueOf(str.charAt(i))).equals("+")
                    || (String.valueOf(str.charAt(i))).equals(".")) {
                //当检测出特殊字符时，添加转义符。
                resStr = resStr + "\\" + String.valueOf(str.charAt(i));
            } else {//非特殊字符直接添加
                resStr += String.valueOf(str.charAt(i));
            }
        }
        return resStr;
    }


}
