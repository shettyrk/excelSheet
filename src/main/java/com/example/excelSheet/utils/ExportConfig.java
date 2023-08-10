package com.example.excelSheet.utils;


import com.example.excelSheet.model.UserActionLog;
import com.example.excelSheet.model.UserActionLogDTO;
import lombok.*;

import java.util.ArrayList;
import java.util.List;

@Data
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class ExportConfig {
    private int sheetIndex;
    private int startRow;
    private Class dataClass;
    private List<CellConfig> cellExportConfigList;
    public static final ExportConfig customerExport;
    static {
        customerExport= new ExportConfig();
        customerExport.setSheetIndex(0);
        customerExport.setStartRow(1);
        customerExport.setDataClass(UserActionLogDTO.class);
        List<CellConfig> userActionLogs = new ArrayList<>();
        userActionLogs.add(new CellConfig(0,"VDMS-ID"));
        userActionLogs.add(new CellConfig(1,"action"));
        userActionLogs.add(new CellConfig(2,"created_timestamp"));
        userActionLogs.add(new CellConfig(3,"email"));
        userActionLogs.add(new CellConfig(4,"message"));
        userActionLogs.add(new CellConfig(5,"status"));
        userActionLogs.add(new CellConfig(6,"type"));
        customerExport.setCellExportConfigList(userActionLogs);
    }
}
