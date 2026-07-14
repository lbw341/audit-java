package com.audit.controller;

import com.audit.service.AuditService;
import com.audit.service.ExcelToHtmlService;
import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.office.OfficeException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.stream.Stream;

/**
 * Web 控制器，处理页面路由、文件上传、审核触发、文件预览和下载。
 *
 * <p>URL 映射：
 * <ul>
 *   <li>GET  /                 — 首页（上传表单 + 审核结果展示）</li>
 *   <li>POST /upload           — AJAX 上传文件（Word 或 Excel）</li>
 *   <li>POST /                 — 执行审核（表单提交）</li>
 *   <li>GET  /open/{name}      — 以原始 MIME 内联提供文件（供 Office 协议调用）</li>
 *   <li>GET  /preview-office/{name} — 高质量预览（Word→PDF / Excel→HTML）</li>
 *   <li>GET  /download/{name}  — 下载文件</li>
 * </ul>
 */
@Controller
public class AuditController {

    @Autowired
    private AuditService auditService;

    /** jodconverter，用于 Word→PDF 转换；如果 LibreOffice 未配置则为 null */
    @Autowired(required = false)
    private DocumentConverter documentConverter;

    @Autowired
    private ExcelToHtmlService excelToHtmlService;

    /** 上传文件存放目录 */
    private final Path uploadDir = Path.of("uploads");

    /** PDF 缓存目录（Word→PDF 结果） */
    private final Path pdfDir = Path.of("pdf_cache");

    public AuditController() throws IOException {
        Files.createDirectories(uploadDir);
        Files.createDirectories(pdfDir);
    }

    /**
     * 文件扩展名到 MIME 类型的映射。
     * 用于 /open/{name} 接口，使浏览器能识别文件类型并调用本地 Office。
     */
    private static final Map<String, MediaType> MIME_MAP = Map.of(
        ".docx", MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        ".doc", MediaType.parseMediaType("application/msword"),
        ".xlsx", MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        ".xlsm", MediaType.parseMediaType("application/vnd.ms-excel.sheet.macroEnabled.12"),
        ".xls", MediaType.parseMediaType("application/vnd.ms-excel")
    );

    /**
     * 首页。
     * 如果是审核后重定向而来（POST /），result 属性已由 audit() 方法设置。
     */
    @GetMapping("/")
    public String index(Model model) {
        model.addAttribute("result", null);
        return "index";
    }

    /**
     * 执行审核。
     * 从 uploads/ 目录查找已上传的 Word 和 Excel 文件，调用 AuditService 比对。
     * 结果保存在原有的 result 属性中，由 Thymeleaf 模板渲染。
     */
    @PostMapping("/")
    public String audit(Model model) {
        // 支持三种 Word 文件名和三种 Excel 文件名
        Path wordDocx = uploadDir.resolve("word_input.docx");
        Path wordDoc = uploadDir.resolve("word_input.doc");
        Path wordPath = Files.exists(wordDocx) ? wordDocx : wordDoc;
        Path excelPath = Stream.of("excel_input.xlsx", "excel_input.xlsm", "excel_input.xls")
            .map(n -> uploadDir.resolve(n)).filter(Files::exists).findFirst().orElse(null);

        if (wordPath != null && excelPath != null && Files.exists(wordPath) && Files.exists(excelPath)) {
            try {
                // 读取上传时保存的原始 excel 文件名，用于构造结果文件名
                Path nameFile = uploadDir.resolve("excel_input.name");
                String originalExcelName = Files.exists(nameFile) ? Files.readString(nameFile).trim() : "excel_input";
                Map<String, Object> result = auditService.audit(wordPath.toString(), excelPath.toString(), originalExcelName);
                model.addAttribute("result", result);
            } catch (Exception e) {
                model.addAttribute("error", e.getMessage());
            }
        }
        return "index";
    }

    /**
     * AJAX 文件上传。
     * 将文件保存为 type_input.ext（如 word_input.docx、excel_input.xlsx），
     * 同时保存原始文件名到 type_input.name，供后续审核时使用。
     *
     * @param file 上传的文件
     * @param type 类型：word 或 excel
     * @return JSON { ok, name, filename, size }
     */
    @PostMapping("/upload")
    @ResponseBody
    public Map<String, Object> upload(@RequestParam("file") MultipartFile file, @RequestParam("type") String type) {
        Map<String, Object> resp = new LinkedHashMap<>();
        if (file.isEmpty()) {
            resp.put("ok", false);
            resp.put("msg", "\u65e0\u6587\u4ef6");  // 无文件
            return resp;
        }
        // 校验扩展名
        String ext = Optional.ofNullable(file.getOriginalFilename())
            .map(f -> f.substring(f.lastIndexOf('.'))).orElse("").toLowerCase();
        if ("word".equals(type)) {
            if (!".docx".equals(ext) && !".doc".equals(ext)) {
                resp.put("ok", false);
                resp.put("msg", "Word\u6587\u4ef6\u9700\u4e3a.doc\u6216.docx\u683c\u5f0f");
                return resp;
            }
        } else if ("excel".equals(type)) {
            if (!".xlsx".equals(ext) && !".xls".equals(ext) && !".xlsm".equals(ext)) {
                resp.put("ok", false);
                resp.put("msg", "Excel\u6587\u4ef6\u9700\u4e3a.xls\u3001.xlsx\u6216.xlsm\u683c\u5f0f");
                return resp;
            }
        } else {
            resp.put("ok", false);
            resp.put("msg", "\u672a\u77e5\u7c7b\u578b");
            return resp;
        }
        try {
            String originalName = file.getOriginalFilename();
            String name = type + "_input" + ext;
            Path dest = uploadDir.resolve(name);
            file.transferTo(dest);
            // 保存原始文件名，供 audit() 读取后传给 AuditService 用于生成结果文件名
            Path nameFile = uploadDir.resolve(type + "_input.name");
            java.nio.file.Files.writeString(nameFile, originalName == null ? name : originalName);
            long size = Files.size(dest);
            String sizeStr = size < 1024 * 1024 ? (size / 1024) + " KB" : String.format("%.1f MB", size / (1024.0 * 1024));
            resp.put("ok", true);
            resp.put("name", originalName);
            resp.put("filename", name);
            resp.put("size", sizeStr);
        } catch (IOException e) {
            resp.put("ok", false);
            resp.put("msg", "\u4e0a\u4f20\u5931\u8d25");
        }
        return resp;
    }

    /**
     * 以内嵌方式提供文件，供 Office 协议（ms-word: / ms-excel:）或 iframe 使用。
     * Content-Disposition 设为 inline，浏览器会尝试用本地 Office 打开。
     */
    @GetMapping("/open/{name}")
    public ResponseEntity<Resource> openFile(@PathVariable String name) {
        Path path = uploadDir.resolve(name);
        if (!Files.exists(path)) {
            return ResponseEntity.notFound().build();
        }
        String ext = name.substring(name.lastIndexOf('.')).toLowerCase();
        MediaType mt = MIME_MAP.getOrDefault(ext, MediaType.APPLICATION_OCTET_STREAM);
        return ResponseEntity.ok()
            .contentType(mt)
            .header(HttpHeaders.CONTENT_DISPOSITION, "inline; filename=\"" + name + "\"")
            .body(new FileSystemResource(path.toFile()));
    }

    /**
     * 高质量预览。
     * <ul>
     *   <li>Word (.doc/.docx) → PDF（通过 jodconverter 调用 LibreOffice）</li>
     *   <li>Excel (.xlsx/.xls/.xlsm) → HTML（通过 ExcelToHtmlService 自定义渲染）</li>
     * </ul>
     */
    @GetMapping("/preview-office/{name}")
    public ResponseEntity<Object> previewOffice(@PathVariable String name) {
        Path srcPath = uploadDir.resolve(name);
        if (!Files.exists(srcPath)) {
            return ResponseEntity.notFound().build();
        }
        String ext = name.substring(name.lastIndexOf('.')).toLowerCase();
        // Excel 文件走 POI 自定义 HTML 渲染
        if (".xlsx".equals(ext) || ".xls".equals(ext) || ".xlsm".equals(ext)) {
            try {
                String html = excelToHtmlService.toHtml(srcPath.toString());
                return ResponseEntity.ok().contentType(MediaType.TEXT_HTML).body(html);
            } catch (IOException e) {
                return ResponseEntity.ok()
                    .contentType(MediaType.TEXT_HTML)
                    .body("<html><body style='font-family:sans-serif;text-align:center;padding:60px;color:#cf1322;'>"
                        + "<h2>渲染失败</h2><p>" + e.getMessage() + "</p>"
                        + "<br><a href='/'>返回首页</a></body></html>");
            }
        }
        // Word 文件走 LibreOffice → PDF
        if (documentConverter == null) {
            return ResponseEntity.ok()
                .contentType(MediaType.TEXT_HTML)
                .body("<html><body style='font-family:sans-serif;text-align:center;padding:60px;color:#666;'>"
                    + "<h2>未安装 LibreOffice</h2>"
                    + "<p>Word 高质量预览需要安装 <a href='https://www.libreoffice.org/download/' target='_blank'>LibreOffice</a>。</p>"
                    + "<br><a href='/'>返回首页</a></body></html>");
        }
        String outName = name.replaceAll("\\.docx?$", ".pdf");
        Path outPath = pdfDir.resolve(outName);
        if (!Files.exists(outPath)) {
            try {
                documentConverter.convert(srcPath.toFile()).to(outPath.toFile()).execute();
            } catch (OfficeException e) {
                return ResponseEntity.ok()
                    .contentType(MediaType.TEXT_HTML)
                    .body("<html><body style='font-family:sans-serif;text-align:center;padding:60px;color:#cf1322;'>"
                        + "<h2>转换失败</h2><p>" + e.getMessage() + "</p>"
                        + "<br><a href='/'>返回首页</a></body></html>");
            }
        }
        return servePdf(outPath, outName);
    }

    /** 提供 PDF 文件给浏览器内联展示 */
    private ResponseEntity<Object> servePdf(Path path, String name) {
        return ResponseEntity.ok()
            .contentType(MediaType.APPLICATION_PDF)
            .header(HttpHeaders.CONTENT_DISPOSITION, "inline; filename=\"" + name + "\"")
            .body(new FileSystemResource(path.toFile()));
    }

    /**
     * 下载文件（附件形式，浏览器不内联展示）。
     * Content-Disposition 设为 attachment，弹出下载对话框。
     * 可通过 ?display_name=xxx 指定下载时显示的文件名（用于结果文件原名含中文的场景）。
     */
    @GetMapping("/download/{name}")
    public ResponseEntity<Resource> downloadFile(@PathVariable String name,
                                                  @RequestParam(value = "display_name", required = false) String displayName) {
        Path path = uploadDir.resolve(name);
        if (!Files.exists(path)) {
            return ResponseEntity.notFound().build();
        }
        String downloadName = (displayName != null && !displayName.isEmpty()) ? displayName : name;
        return ResponseEntity.ok()
            .contentType(MediaType.APPLICATION_OCTET_STREAM)
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + downloadName + "\"")
            .body(new FileSystemResource(path.toFile()));
    }
}
