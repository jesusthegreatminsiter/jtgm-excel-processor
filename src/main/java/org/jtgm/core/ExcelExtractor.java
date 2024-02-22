package org.jtgm.core;

import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;

public interface ExcelExtractor {
    void extract(MultipartFile file);
}
