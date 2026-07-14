package com.audit.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

@Service
public class ExcelToHtmlService {

    public String toHtml(String filePath) throws IOException {
        StringBuilder sb = new StringBuilder();
        sb.append("<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"UTF-8\">");
        sb.append("<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">");
        sb.append("<title>").append(escapeHtml(new java.io.File(filePath).getName())).append("</title>");
        sb.append("<style>");
        sb.append("*{margin:0;padding:0;box-sizing:border-box;}");
        sb.append("body{font-family:-apple-system,'Microsoft YaHei',sans-serif;background:#f0f2f5;padding:20px;}");
        sb.append(".sheet-tabs{display:flex;gap:2px;margin-bottom:12px;overflow-x:auto;}");
        sb.append(".sheet-tab{padding:8px 20px;border-radius:6px 6px 0 0;font-size:13px;cursor:pointer;background:#e8e8e8;color:#555;white-space:nowrap;border:none;}");
        sb.append(".sheet-tab.active{background:#fff;color:#1677ff;font-weight:600;box-shadow:0 -1px 3px rgba(0,0,0,.06);}");
        sb.append(".sheet-content{display:none;}");
        sb.append(".sheet-content.active{display:block;}");
        sb.append(".excel-table{border-collapse:collapse;width:auto;font-size:13px;background:#fff;box-shadow:0 1px 6px rgba(0,0,0,.06);border-radius:8px;}");
        sb.append(".excel-table td{border:1px solid #d0d5dd;padding:4px 8px;white-space:nowrap;min-width:60px;max-width:400px;overflow:hidden;text-overflow:ellipsis;}");
        sb.append(".excel-table thead td{background:#f7f8fa;font-weight:600;border-bottom:2px solid #b0b5bd;position:sticky;top:0;z-index:1;}");
        sb.append(".green-bg{background-color:#92D050 !important;}");
        sb.append(".yellow-bg{background-color:#FFFF00 !important;}");
        sb.append(".wrap-container{overflow:auto;max-height:80vh;}");
        sb.append("</style></head><body>");
        sb.append("<div class=\"sheet-tabs\" id=\"tabs\"></div>");
        sb.append("<div class=\"wrap-container\" id=\"sheets\">");

        try (Workbook wb = WorkbookFactory.create(new FileInputStream(filePath))) {
            int sheetCount = wb.getNumberOfSheets();
            for (int i = 0; i < sheetCount; i++) {
                Sheet sheet = wb.getSheetAt(i);
                String sheetName = wb.getSheetName(i);

                int firstRow = sheet.getFirstRowNum();
                int lastRow = sheet.getLastRowNum();
                if (lastRow < 0) continue;

                int maxCol = 0;
                for (int r = firstRow; r <= lastRow; r++) {
                    Row row = sheet.getRow(r);
                    if (row != null && row.getLastCellNum() > maxCol) maxCol = row.getLastCellNum();
                }
                if (maxCol == 0) continue;

                int headerRowCnt = 0;
                Row hRow = sheet.getRow(firstRow);
                if (hRow != null) {
                    for (int c = 0; c < maxCol; c++) {
                        Cell cell = hRow.getCell(c);
                        if (cell != null && !getCellValue(cell).isEmpty()) { headerRowCnt = 1; break; }
                    }
                }

                List<CellRangeAddress> merges = sheet.getMergedRegions();

                sb.append("<div class=\"sheet-content").append(i == 0 ? " active" : "").append("\" data-sheet=\"").append(escapeHtml(sheetName)).append("\">");
                sb.append("<table class=\"excel-table\" id=\"sheet-").append(i).append("\">");
                if (headerRowCnt > 0) sb.append("<thead>");
                sb.append("<tbody>");

                for (int r = firstRow; r <= lastRow; r++) {
                    Row row = sheet.getRow(r);
                    sb.append("<tr>");

                    Set<Integer> skipCols = new HashSet<>();
                    Map<Integer, Integer[]> mergeAttrs = new HashMap<>();

                    for (CellRangeAddress m : merges) {
                        if (m.getFirstRow() <= r && m.getLastRow() >= r) {
                            if (r > m.getFirstRow() && r <= m.getLastRow()) {
                                for (int c = m.getFirstColumn(); c <= m.getLastColumn(); c++) {
                                    skipCols.add(c);
                                }
                            }
                        }
                    }

                    for (CellRangeAddress m : merges) {
                        if (m.getFirstRow() == r) {
                            int colspan = m.getLastColumn() - m.getFirstColumn() + 1;
                            int rowspan = m.getLastRow() - m.getFirstRow() + 1;
                            for (int c = m.getFirstColumn() + 1; c <= m.getLastColumn(); c++) {
                                skipCols.add(c);
                            }
                            mergeAttrs.put(m.getFirstColumn(), new Integer[]{colspan, rowspan});
                        }
                    }

                    if (row == null) {
                        for (int c = 0; c < maxCol; c++) {
                            if (skipCols.contains(c)) continue;
                            sb.append("<td>&nbsp;</td>");
                        }
                    } else {
                        int[] colWidths = new int[maxCol];
                        for (int c = 0; c < maxCol; c++) {
                            int w = sheet.getColumnWidth(c);
                            colWidths[c] = Math.max(60, Math.min(w * 7 / 256, 400));
                        }

                        for (int c = 0; c < maxCol; c++) {
                            if (skipCols.contains(c)) continue;

                            Cell cell = row.getCell(c);
                            String val = "";
                            String cls = "";
                            String extra = "";

                            if (mergeAttrs.containsKey(c)) {
                                Integer[] a = mergeAttrs.get(c);
                                extra = " colspan=\"" + a[0] + "\"";
                                if (a[1] > 1) extra += " rowspan=\"" + a[1] + "\"";
                            }

                            if (cell != null) {
                                val = escapeHtml(getCellValue(cell));
                                String bg = getCellBgHex(cell, wb);
                                if (bg != null) {
                                    if ("#92D050".equalsIgnoreCase(bg)) {
                                        cls = "green-bg";
                                    } else if ("#FFFF00".equalsIgnoreCase(bg)) {
                                        cls = "yellow-bg";
                                    } else if (!"#FFFFFF".equalsIgnoreCase(bg) && !"#00000000".equalsIgnoreCase(bg)) {
                                        extra += " style=\"background-color:" + bg + "\"";
                                    }
                                }
                            }

                            sb.append("<td");
                            if (!cls.isEmpty()) sb.append(" class=\"").append(cls).append("\"");
                            sb.append(extra);
                            sb.append(" style=\"min-width:").append(colWidths[c]).append("px\"");
                            sb.append(">").append(val.isEmpty() ? "&nbsp;" : val).append("</td>");
                        }
                    }
                    sb.append("</tr>");
                }

                sb.append("</tbody>");
                if (headerRowCnt > 0) sb.append("</thead>");
                sb.append("</table></div>");
            }
        }

        sb.append("</div>");
        sb.append("<script>");
        sb.append("(function(){");
        sb.append("var tabs=document.getElementById('tabs');");
        sb.append("var sheets=document.getElementById('sheets').children;");
        sb.append("for(var i=0;i<sheets.length;i++){");
        sb.append("(function(idx){");
        sb.append("var tab=document.createElement('button');");
        sb.append("tab.className='sheet-tab'+(idx===0?' active':'');");
        sb.append("tab.textContent=sheets[idx].getAttribute('data-sheet');");
        sb.append("tab.onclick=function(){");
        sb.append("document.querySelectorAll('.sheet-tab').forEach(function(t){t.classList.remove('active');});");
        sb.append("document.querySelectorAll('.sheet-content').forEach(function(s){s.classList.remove('active');});");
        sb.append("tab.classList.add('active');");
        sb.append("sheets[idx].classList.add('active');");
        sb.append("};");
        sb.append("tabs.appendChild(tab);");
        sb.append("})(i);");
        sb.append("}");
        sb.append("})();");
        sb.append("</script>");
        sb.append("</body></html>");
        return sb.toString();
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                double v = cell.getNumericCellValue();
                if (v == Math.floor(v) && !Double.isInfinite(v)) yield String.valueOf((long) v);
                else yield String.valueOf(v);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                try { yield String.valueOf((int) cell.getNumericCellValue()); }
                catch (Exception e) { try { yield cell.getStringCellValue(); } catch (Exception e2) { yield ""; } }
            }
            default -> "";
        };
    }

    private String getCellBgHex(Cell cell, Workbook wb) {
        CellStyle cs = cell.getCellStyle();
        // 未设置填充模式时跳过（默认 NO_FILL），避免 .xls 中索引 0 映射为黑色
        if (cs.getFillPattern() == FillPatternType.NO_FILL) {
            return null;
        }
        if (cs instanceof XSSFCellStyle xcs) {
            if (xcs.getFillForegroundColorColor() != null) {
                String hex = xcs.getFillForegroundColorColor().getARGBHex();
                if (hex != null) return hex.length() > 6 ? "#" + hex.substring(2) : "#" + hex;
            }
        } else if (cs instanceof HSSFCellStyle hcs && wb instanceof HSSFWorkbook hwb) {
            short idx = hcs.getFillForegroundColor();
            // 索引 0 表示"自动/无填充"，忽略
            if (idx == 0) return null;
            // 使用工作簿的自定义调色板，而非静态 HSSFColor.getIndexHash()
            // 因为 .xls 文件中可能使用了自定义颜色，静态 map 无法覆盖
            HSSFPalette palette = hwb.getCustomPalette();
            HSSFColor color = palette.getColor(idx);
            if (color != null) {
                short[] rgb = color.getTriplet();
                return String.format("#%02X%02X%02X", rgb[0], rgb[1], rgb[2]);
            }
        }
        return null;
    }

    private String escapeHtml(String s) {
        if (s == null) return "";
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                .replace("\"", "&quot;").replace("'", "&#39;");
    }
}
