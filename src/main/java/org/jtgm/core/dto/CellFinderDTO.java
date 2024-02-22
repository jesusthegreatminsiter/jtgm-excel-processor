package org.jtgm.core.dto;

import lombok.*;
import org.apache.poi.ss.usermodel.Row;

import java.util.HashMap;
import java.util.Map;

@Builder
@Getter
@Setter
@AllArgsConstructor
public class CellFinderDTO {
    private Row row;
    Map<String, Integer> foundHeaderMap;

    public static CellFinderDTO buildCellFinder(HashMap<String, Integer> headers, Row row) {
        return builder()
                .row(row)
                .foundHeaderMap(headers)
                .build();
    }
}
