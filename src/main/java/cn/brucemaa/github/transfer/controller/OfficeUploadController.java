package cn.brucemaa.github.transfer.controller;

import cn.brucemaa.github.transfer.transfer.ExcelComponent;
import cn.brucemaa.github.transfer.transfer.PowerPointComponent;
import cn.brucemaa.github.transfer.transfer.WordComponent;
import org.apache.commons.io.FilenameUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;

/**
 * projectName:office2pdf-transfer-demo
 * cn.brucemaa.github.transfer.controller
 *
 * @author Bruce Maa
 * @since 2019-01-23.14:36
 */
@RestController
public class OfficeUploadController {

    private final Logger logger = LoggerFactory.getLogger(this.getClass());


    /**
     * 上传文档
     * @return 转换结果
     */
    @PostMapping(value = "/upload")
    public ResponseEntity upload(@RequestParam("file") MultipartFile file) {

        String tmpDir = System.getProperty("user.dir") + File.separator;
        String localFileName = tmpDir + file.getOriginalFilename();
        File localFile = new File(localFileName);

        // 获取文档类型
        String fileExt = FilenameUtils.getExtension(file.getOriginalFilename());
        String fileName = FilenameUtils.getName(file.getOriginalFilename());
        String localPdfFileName = tmpDir + fileName + ".pdf";

        try {
            file.transferTo(localFile);
            if ("doc".equalsIgnoreCase(fileExt) || "docx".equalsIgnoreCase(fileExt)) {
                WordComponent wordComponent = new WordComponent();
                wordComponent.saveAsPdf(localFileName, localPdfFileName);
            } else if ("xls".equalsIgnoreCase(fileExt) || "xlsx".equalsIgnoreCase(fileExt)) {
                ExcelComponent excelComponent = new ExcelComponent();
                excelComponent.saveAsPdf(localFileName, localPdfFileName);
            } else if ("ppt".equalsIgnoreCase(fileExt) || "pptx".equalsIgnoreCase(fileExt)) {
                PowerPointComponent powerPointComponent = new PowerPointComponent();
                powerPointComponent.saveAsPdf(localFileName, localPdfFileName);
            }
            return ResponseEntity.ok("转换成功");
        } catch (Exception e) {
            logger.error("file transfer pdf error: {}", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("转换失败");
        }
    }
}
