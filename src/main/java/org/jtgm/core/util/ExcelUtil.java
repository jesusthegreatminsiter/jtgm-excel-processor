package org.jtgm.core.util;

import org.apache.poi.ss.usermodel.Sheet;

public interface ExcelUtil {
    void execute(Sheet sheet, String mgroupName);
}
