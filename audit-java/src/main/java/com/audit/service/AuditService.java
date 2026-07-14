package com.audit.service;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.springframework.stereotype.Service;

import java.io.*;
import java.nio.file.Path;
import java.util.*;

/**
 * 核心业务逻辑：读取 Word 和 Excel，按三级模块名称比对，标色后生成结果文件。
 *
 * <p>Word 支持 .docx（POI 直接读）和 .doc（需通过 LibreOffice 转成 .docx）。
 * Excel 支持 .xlsx（POI 直读）、.xlsm（POI 直读）和 .xls（LibreOffice 转 .xlsx）。
 *
 * <p>比对逻辑：
 * <ul>
 *   <li>绿色 = Word 和 Excel 中都存在的模块（匹配成功）</li>
 *   <li>黄色 = Excel 中有但 Word 中没有的模块（额外）</li>
 *   <li>无填充 = Word 中有但 Excel 中没有的模块（缺失）——不缺标注，见缺失列表</li>
 * </ul>
 */
@Service
public class AuditService {

    /**
     * LibreOffice 可执行文件路径。
     * 用于将 .doc 转为 .docx、.xls 转为 .xlsx，因为 POI 的 XWPFDocument / XSSFWorkbook
     * 不兼容旧格式（.doc 是 OLE 容器，.xls 是 BIFF 格式）。
     */
    private static final String LIBRE_OFFICE = "C:\\Program Files\\LibreOffice\\program\\soffice.exe";

    /**
     * 在 Excel 表头行中搜索的目标列名。
     * 只有包含此列的 sheet 才会参与审核，同时会用该列去定位一、二级模块列。
     */
    private static final String TARGET_HEADER = "\u4e09\u7ea7\u6a21\u5757";

    /**
     * 记录一个 sheet 中找到的目标表头位置及一、二级模块列位置。
     *
     * @param sheet     所在的 XSSFSheet 对象
     * @param headerRow 表头所在的行号（0-based）
     * @param col       三级模块所在的列号（1-based）
     * @param name      sheet 名称
     * @param level1Col 一级模块列号（1-based），未找到时为 lastCol - 2
     * @param level2Col 二级模块列号（1-based），未找到时为 lastCol - 1
     */
    record SheetInfo(XSSFSheet sheet, int headerRow, int col, String name, int level1Col, int level2Col) {}

    /**
     * 从 Excel 中读出的单条三级模块记录，保留其行号、列号和所属层级信息。
     *
     * @param row    数据所在行号（0-based），用于后续标色时定位
     * @param col    三级模块列号（1-based）
     * @param level1 同行的"一级模块"值，可能为空
     * @param level2 同行的"二级模块"值，可能为空
     * @param level3 三级模块名称（去空白后的文本）
     * @param sheet  所属 sheet 名称
     */
    record ExcelModule(int row, int col, String level1, String level2, String level3, String sheet) {}

    /**
     * 执行三级模块审核。
     *
     * <p>处理流程：
     * <ol>
     *   <li>从 Word 文档的第一个表格第4列提取所有三级模块</li>
     *   <li>非 .xlsx 的 Excel 通过 LibreOffice 转为临时 .xlsx</li>
     *   <li>在所有 sheet 中搜索"三级模块"表头，读取其下方的所有数据行</li>
     *   <li>用集合运算求出匹配/Word独有/Excel独有三组</li>
     *   <li>在结果 Excel 中逐行标色（绿色=匹配，黄色=额外）</li>
     *   <li>以用户原始文件名 + "审核结果.xlsx" 保存到上传目录</li>
     * </ol>
     *
     * @param wordPath          Word 文件的完整路径
     * @param excelPath         Excel 文件的完整路径
     * @param originalExcelName 用户上传时的原始文件名（用于构造结果文件名）
     * @return 包含统计数据和差异列表的 Map，直接传给 Thymeleaf 模板渲染
     * @throws IOException 文件读写或 LibreOffice 转换失败时抛出
     */
    public Map<String, Object> audit(String wordPath, String excelPath, String originalExcelName) throws IOException {
        // 放宽 ZIP 膨胀比限制，防止某些 docx/xlsx 因压缩比异常而报错
        ZipSecureFile.setMinInflateRatio(0);
        // 记住原始路径，后续可能被临时路径覆盖，但输出文件必须写到上传目录
        String originalExcelPath = excelPath;

        // ==================== 1. 从 Word 取三级模块列表 ====================
        List<String> wordModules = readWordModules(wordPath);
        Set<String> wordSet = new LinkedHashSet<>(wordModules);

        // ==================== 2. 非 .xlsx 的 Excel 用 LibreOffice 转成 .xlsx ====================
        Path tempDir = null;
        if (!excelPath.toLowerCase().endsWith(".xlsx")) {
            // 创建临时目录存放转换结果，finally 中清理
            tempDir = java.nio.file.Files.createTempDirectory("xlsconv");
            String baseName = new File(excelPath).getName();
            if (baseName.contains(".")) baseName = baseName.substring(0, baseName.lastIndexOf('.'));
            Path convertedPath = tempDir.resolve(baseName + ".xlsx");
            try {
                // 调用 soffice.exe --headless --convert-to xlsx 进行后台转换
                ProcessBuilder pb = new ProcessBuilder(LIBRE_OFFICE, "--headless", "--convert-to", "xlsx",
                    "--outdir", tempDir.toString(), excelPath);
                pb.redirectErrorStream(true);
                Process p = pb.start();
                p.waitFor();
                if (convertedPath.toFile().exists()) {
                    excelPath = convertedPath.toString();
                } else {
                    throw new IOException("LibreOffice 转换 Excel 失败");
                }
            } catch (InterruptedException e) {
                throw new IOException("LibreOffice 转换被中断", e);
            }
        }

        // ==================== 3. 打开 Excel 工作簿（此时 excelPath 一定是 .xlsx） ====================
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(excelPath));
        try {
            // 搜索所有包含"三级模块"表头的 sheet
            List<SheetInfo> sheets = findSheets(wb);
            if (sheets.isEmpty()) {
                throw new IllegalArgumentException("\u672a\u627e\u5230\u201c\u4e09\u7ea7\u6a21\u5757\u201d\u5217");
            }

            // ==================== 4. 读取每个 sheet 的所有三级模块 ====================
            List<ExcelModule> allExcel = new ArrayList<>();
            // sheetData 保留每个 sheet 的原始数据，供后续标色时按 sheet 逐行处理
            List<Object[]> sheetData = new ArrayList<>();

            for (SheetInfo si : sheets) {
                List<ExcelModule> mods = new ArrayList<>();
                // 从表头下一行开始读到最后一行
                for (int r = si.headerRow + 1; r <= si.sheet.getLastRowNum(); r++) {
                    Row row = si.sheet.getRow(r);
                    if (row == null) continue;
                    Cell c = row.getCell(si.col - 1);   // col 是 1-based，POI 需要 0-based
                    if (c == null) continue;
                    String l3 = getCellValue(c);
                    if (l3.isEmpty()) continue;
                    // 读取一、二级模块列（如果存在），level1Col/level2Col 为 0 或负数时表示不存在
                    String l1 = si.level1Col > 0 ? getCellValue(row.getCell(si.level1Col - 1)) : "";
                    String l2 = si.level2Col > 0 ? getCellValue(row.getCell(si.level2Col - 1)) : "";
                    mods.add(new ExcelModule(r + 1, si.col, l1, l2, l3, si.name));
                }
                allExcel.addAll(mods);
                sheetData.add(new Object[]{si, mods});
            }

            // ==================== 5. 集合运算：匹配、Word 独有、Excel 独有 ====================
            Set<String> excelSet = new LinkedHashSet<>();
            for (ExcelModule m : allExcel) excelSet.add(m.level3);

            // matched  = wordSet ∩ excelSet
            Set<String> matched = new LinkedHashSet<>(wordSet);
            matched.retainAll(excelSet);

            // onlyInWord  = wordSet - excelSet
            Set<String> onlyInWord = new LinkedHashSet<>(wordSet);
            onlyInWord.removeAll(excelSet);

            // onlyInExcel = excelSet - wordSet
            Set<String> onlyInExcel = new LinkedHashSet<>(excelSet);
            onlyInExcel.removeAll(wordSet);

            // ==================== 6. 定义标记颜色 ====================
            // 绿色 (#92D050)：Word 和 Excel 均存在，匹配成功
            XSSFColor greenColor = new XSSFColor(new byte[]{(byte)0x92, (byte)0xD0, 0x50}, new DefaultIndexedColorMap());
            // 黄色 (#FFFF00)：Excel 中有但 Word 中没有，额外项
            XSSFColor yellowColor = new XSSFColor(new byte[]{(byte)0xFF, (byte)0xFF, 0x00}, new DefaultIndexedColorMap());

            int greenCount = 0, yellowCount = 0;

            // ==================== 7. 逐行标色 ====================
            // 注：复用两个 CellStyle 而非每行创建，避免超出 Excel 64000 样式上限
            XSSFCellStyle greenStyle = wb.createCellStyle();
            greenStyle.setFillForegroundColor(greenColor);
            greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            XSSFCellStyle yellowStyle = wb.createCellStyle();
            yellowStyle.setFillForegroundColor(yellowColor);
            yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            for (Object[] sd : sheetData) {
                SheetInfo si = (SheetInfo) sd[0];
                @SuppressWarnings("unchecked")
                List<ExcelModule> mods = (List<ExcelModule>) sd[1];
                for (ExcelModule em : mods) {
                    Row row = si.sheet.getRow(em.row - 1);   // em.row 存储时是 1-based
                    if (row == null) continue;
                    Cell cell = row.getCell(em.col - 1);     // em.col 同样是 1-based
                    if (cell == null) continue;
                    if (matched.contains(em.level3)) {
                        cell.setCellStyle(greenStyle);
                        greenCount++;
                    } else if (onlyInExcel.contains(em.level3)) {
                        cell.setCellStyle(yellowStyle);
                        yellowCount++;
                    }
                    // 缺失项（onlyInWord）不标注，保持原样
                }
            }

            // ==================== 8. 生成结果文件名 ====================
            // 内部使用纯 ASCII 文件名（避免 ms-office: 协议不支持中文），用于 URL 链接
            String outName = "result.xlsx";
            // display_name 用于下载时的 Content-Disposition（用户看到的原始文件名 + "审核结果"）
            String displayBase = originalExcelName;
            if (displayBase.contains(".")) displayBase = displayBase.substring(0, displayBase.lastIndexOf('.'));
            String displayName = displayBase + "\u5ba1\u6838\u7ed3\u679c.xlsx";
            // 输出到上传目录，而非临时目录（否则会被 finally 清理掉）
            String outPath = new File(originalExcelPath).getParent() + File.separator + outName;

            // ==================== 9. 保存标色后的 Excel ====================
            try (FileOutputStream fos = new FileOutputStream(outPath)) {
                wb.write(fos);
            }

            // ==================== 10. 整理缺失/额外列表（用于页面展示） ====================
            List<String> missingList = new ArrayList<>(onlyInWord);

            List<Map<String, String>> extraList = new ArrayList<>();
            for (ExcelModule em : allExcel) {
                if (onlyInExcel.contains(em.level3)) {
                    Map<String, String> m = new LinkedHashMap<>();
                    m.put("level1", em.level1);
                    m.put("level2", em.level2);
                    m.put("level3", em.level3);
                    m.put("sheet", em.sheet);
                    extraList.add(m);
                }
            }

            // ==================== 11. 组装结果对象 ====================
            int totalWord = wordSet.size();
            int totalExcel = excelSet.size();
            int matchCount = matched.size();
            int pct = totalWord > 0 ? (int) Math.round(matchCount * 100.0 / totalWord) : 0;

            Map<String, Object> result = new LinkedHashMap<>();
            result.put("total_word", totalWord);
            result.put("total_excel", totalExcel);
            result.put("match_count", matchCount);
            result.put("pct", pct);                     // 匹配率（百分制整数）
            result.put("green", greenCount);             // 实际标绿行数
            result.put("yellow", yellowCount);           // 实际标黄行数
            result.put("missing_word", missingList);     // 缺失列表（纯文本）
            result.put("extra_excel", extraList);        // 额外列表（含层级和 sheet 名）
            result.put("out_name", outName);             // 内部 ASCII 文件名（用于预览/下载链接）
            result.put("display_name", displayName);     // 用户原始文件名+"审核结果"（用于下载弹窗显示）
            result.put("word_name", new File(wordPath).getName());
            result.put("excel_name", originalExcelName);
            return result;

        } finally {
            // 无论如何都要关闭工作簿并清理 LibreOffice 转换产生的临时目录
            wb.close();
            if (tempDir != null) {
                try {
                    java.nio.file.Files.walk(tempDir)
                        .sorted(java.util.Comparator.reverseOrder())
                        .map(java.nio.file.Path::toFile)
                        .forEach(File::delete);
                } catch (Exception ignored) {
                    // 临时目录清理失败不影响主流程
                }
            }
        }
    }

    /**
     * 从 Word 文档的第一个表格第4列提取三级模块名称。
     *
     * <p>.docx 直接用 POI 的 XWPFDocument 读取；
     * .doc 先通过 LibreOffice 转换成 .docx 到临时目录，再用相同方式读取。
     *
     * @param wordPath Word 文件路径（.doc 或 .docx）
     * @return 三级模块名称列表，按 Word 文档中的原始顺序
     * @throws IOException 读取失败或 LibreOffice 转换失败时抛出
     */
    private List<String> readWordModules(String wordPath) throws IOException {
        String path = wordPath;
        Path tempDir = null;

        // .doc 格式需要先转成 .docx
        if (wordPath.toLowerCase().endsWith(".doc")) {
            tempDir = java.nio.file.Files.createTempDirectory("docconv");
            String baseName = new File(wordPath).getName();
            if (baseName.contains(".")) baseName = baseName.substring(0, baseName.lastIndexOf('.'));
            Path outPath = tempDir.resolve(baseName + ".docx");
            try {
                ProcessBuilder pb = new ProcessBuilder(LIBRE_OFFICE, "--headless", "--convert-to", "docx",
                    "--outdir", tempDir.toString(), wordPath);
                pb.redirectErrorStream(true);
                Process p = pb.start();
                p.waitFor();
                if (outPath.toFile().exists()) {
                    path = outPath.toString();
                } else {
                    throw new IOException("LibreOffice 转换 .doc 失败");
                }
            } catch (Exception e) {
                throw new IOException("LibreOffice 转换 .doc 失败: " + e.getMessage());
            }
        }

        try {
            return readDocxModules(path);
        } finally {
            // 清理 .doc 转换产生的临时 .docx 文件
            if (tempDir != null) {
                try {
                    java.nio.file.Files.walk(tempDir)
                        .sorted(java.util.Comparator.reverseOrder())
                        .map(java.nio.file.Path::toFile)
                        .forEach(File::delete);
                } catch (Exception ignored) {
                }
            }
        }
    }

    /**
     * 读取 .docx 中的第一个表格，提取第4列（索引 3）所有非空文本。
     * 跳过表头行（索引 0），从第 2 行（索引 1）开始遍历。
     *
     * @param wordPath .docx 文件路径
     * @return 三级模块名称列表
     * @throws IOException POI 解析失败时抛出
     */
    private List<String> readDocxModules(String wordPath) throws IOException {
        List<String> modules = new ArrayList<>();
        try (XWPFDocument doc = new XWPFDocument(new FileInputStream(wordPath))) {
            List<XWPFTable> tables = doc.getTables();
            if (tables.isEmpty()) return modules;
            XWPFTable table = tables.get(0);
            // 跳过表头（i=0），从数据行（i=1）开始
            for (int i = 1; i < table.getRows().size(); i++) {
                XWPFTableRow row = table.getRow(i);
                List<XWPFTableCell> cells = row.getTableCells();
                // 第4列（下标 3）为三级模块
                if (cells.size() >= 4) {
                    String text = cells.get(3).getText().trim();
                    if (!text.isEmpty()) modules.add(text);
                }
            }
        }
        return modules;
    }

    /**
     * 在所有 sheet 的前10行中搜索"三级模块"表头，返回找到的 SheetInfo 列表。
     *
     * <p>搜索策略：
     * <ul>
     *   <li>每个 sheet 只找第一次出现的位置</li>
     *   <li>只扫描前10行以提升性能</li>
     *   <li>找到三级模块后，在同一行中继续查找"一级模块"和"二级模块"表头</li>
     *   <li>如果找不到一/二级模块，按"紧邻左边两列"的规则推测（lastCol-2, lastCol-1）</li>
     * </ul>
     *
     * @param wb 已打开的 XSSFWorkbook
     * @return 所有包含"三级模块"的 sheet 信息
     */
    private List<SheetInfo> findSheets(XSSFWorkbook wb) {
        List<SheetInfo> results = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            XSSFSheet sheet = wb.getSheetAt(i);
            String sheetName = wb.getSheetName(i);
            int lastCol = -1;
            // 扫描前10行，找到第一个"三级模块"表头
            for (int r = 0; r < Math.min(sheet.getLastRowNum() + 1, 10); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    Cell cell = row.getCell(c);
                    if (cell != null && TARGET_HEADER.equals(getCellValue(cell))) {
                        lastCol = c + 1;   // 转为 1-based
                        break;
                    }
                }
                if (lastCol > 0) {
                    // 找到三级模块列，在同一行中搜索一、二级模块表头
                    int level1Col = -1, level2Col = -1;
                    Row hRow = sheet.getRow(r);
                    if (hRow != null) {
                        for (int c = 0; c < hRow.getLastCellNum(); c++) {
                            String v = getCellValue(hRow.getCell(c));
                            if ("\u4e00\u7ea7\u6a21\u5757".equals(v)) level1Col = c + 1;
                            else if ("\u4e8c\u7ea7\u6a21\u5757".equals(v)) level2Col = c + 1;
                        }
                    }
                    // 如果没找到一/二级模块，根据常规表格布局推测列位置
                    // 常见布局：... | 一级模块 | 二级模块 | 三级模块
                    if (level1Col < 0) level1Col = lastCol - 2;
                    if (level2Col < 0) level2Col = lastCol - 1;
                    results.add(new SheetInfo(sheet, r, lastCol, sheetName, level1Col, level2Col));
                    break;   // 一个 sheet 只处理第一个匹配的表头
                }
            }
        }
        return results;
    }

    /**
     * 获取单元格的字符串值，统一处理各种数据类型。
     *
     * <p>处理规则：
     * <ul>
     *   <li>STRING：直接返回去空白文本</li>
     *   <li>NUMERIC：整数返回不带小数点的字符串（如 123），小数原样返回</li>
     *   <li>BOOLEAN：返回 "true" 或 "false"</li>
     *   <li>FORMULA / BLANK / ERROR：返回空字符串（不处理公式求值）</li>
     * </ul>
     *
     * @param cell 单元格对象，可能为 null
     * @return 单元格的文本值，null 返回空字符串
     */
    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> {
                double v = cell.getNumericCellValue();
                // 整数不显示小数点（如 1.0 → "1"）
                if (v == Math.floor(v) && !Double.isInfinite(v)) {
                    yield String.valueOf((long) v);
                }
                yield String.valueOf(v);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            // FORMULA / BLANK / ERROR 均视为空
            default -> "";
        };
    }
}
