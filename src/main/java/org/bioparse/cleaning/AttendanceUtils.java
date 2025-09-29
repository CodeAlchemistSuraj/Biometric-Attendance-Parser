package org.bioparse.cleaning;

import java.util.List;

public class AttendanceUtils {
    public static String safeGet(List<String> list, int i) {
        return (list != null && i < list.size() && list.get(i) != null) ? list.get(i).trim() : "";
    }
}