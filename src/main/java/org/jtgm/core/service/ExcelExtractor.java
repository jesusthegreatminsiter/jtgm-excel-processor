package org.jtgm.core.service;

import org.springframework.web.multipart.MultipartFile;

public interface ExcelExtractor {
    void extract(MultipartFile file);
}
