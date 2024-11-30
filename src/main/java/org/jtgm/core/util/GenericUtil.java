package org.jtgm.core.util;

import java.io.File;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Calendar;
import java.util.Date;

public class GenericUtil {
    private static final String RESOURCE_NAME = "/file/Transactional.xlsx";

    public static String removeSpaces(String toFormat){
        return toFormat.replaceAll("\\s", "");
    }

    public static void copyFile(File out) throws Exception {
        Path path = FileSystems.getDefault().getPath(out.getPath());
        Files.copy(ExcelUtil.class.getResourceAsStream(RESOURCE_NAME), path, StandardCopyOption.REPLACE_EXISTING);
    }

    public static int computeWeekNumber(Date date) {
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        return calendar.get(Calendar.WEEK_OF_YEAR);
    }

    public static Date getFridayOfWeek(Date date) {
        Calendar calendar = Calendar.getInstance();
        calendar.set(Calendar.WEEK_OF_YEAR, computeWeekNumber(date));
        calendar.set(Calendar.DAY_OF_WEEK, Calendar.FRIDAY);

        return calendar.getTime();
    }

}
