package org.jtgm.core.util;

import org.apache.poi.ss.usermodel.Sheet;

public interface ExcelUtil {
    void generateWorkBook(Sheet sheet, String mgroupName);
}
