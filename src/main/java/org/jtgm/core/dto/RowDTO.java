package org.jtgm.core.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

@Builder
@Getter
@Setter
@AllArgsConstructor
public class RowDTO {
    private String fullName;
    private int weekNumber;
    private String mgroup;
}
