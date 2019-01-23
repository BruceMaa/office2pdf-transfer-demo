package cn.brucemaa.github.transfer.transfer;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComFailException;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * projectName:office2pdf-transfer-demo
 * cn.brucemaa.github.transfer.transfer
 *
 * @author Bruce Maa
 * @since 2019-01-23.14:36
 */
public class PowerPointComponent {
    private final Logger logger = LoggerFactory.getLogger(this.getClass());

    private ActiveXComponent objPpt;
    private Dispatch presentation;

    private String filename;

    /**
     * 打开ppt幻灯片
     *
     * @param filename ppt幻灯片路径
     */
    private void open(String filename) {
        this.filename = filename;

        // 初始化com线程
        ComThread.InitMTA(true);

        objPpt = new ActiveXComponent("PowerPoint.Application");
        // 禁用宏
        objPpt.setProperty("AutomationSecurity", new Variant(3));
        objPpt.setProperty("DisplayAlerts", new Variant(0));
        Dispatch presentations = objPpt.getProperty("Presentations").toDispatch();
        try {
            presentation = Dispatch.call(presentations, "Open", filename,
                    // 是否只读
                    new Variant(true),
                    // Untitled指定文件是否有标题
                    new Variant(true),
                    // WithWindow指定文件是否可见
                    new Variant(false)
            ).toDispatch();
        } catch (ComFailException e) {
            logger.error("ppt filename: {}, open error: {}", filename, e);
            throw e;
        } catch (Exception e) {
            logger.error("ppt filename: {}, open error: {}", filename, e);
            throw e;
        }
    }

    /**
     * 将ppt幻灯片另存为pdf文件
     *
     * @param pdfFilename pdf文件地址
     */
    private void saveAsPdf(String pdfFilename) {
        try {
            // 另存储为PDF文件
            Variant variant = Dispatch.call(presentation, "SaveAs", pdfFilename, 32);
            logger.info("ppt filename: {}, pdfFilename: {}, saveAs: {}", filename, pdfFilename, variant);
        } catch (Exception e) {
            logger.error("ppt filename: {}, SaveAs error: {}", filename, e);
            throw e;
        }
    }

    /**
     * 将ppt幻灯片关闭
     */
    private void close() {
        try {
            Variant variant = Dispatch.call(presentation, "Close", false);
            logger.info("ppt filename: {}, Close: {}", filename, variant);
            variant = objPpt.invoke("Quit", new Variant[0]);
            logger.info("ppt filename: {}, Quit: {}", filename, variant);
            variant.safeRelease();
        } catch (Exception e) {
            logger.error("ppt filename: {}, Close error: {}", filename, e);
            throw e;
        } finally {
            // 释放com线程
            ComThread.Release();
        }
    }

    public void saveAsPdf(String sourceFilename, String pdfFilename) {
        try {
            open(sourceFilename);
            saveAsPdf(pdfFilename);
        } catch (ComFailException e) {
            logger.error("ppt filename: {}, open error: {}", filename, e);
            throw e;
        } catch (Exception e) {
            logger.error("PowerPointComponent saveAsPdf error: {}", e);
            throw e;
        } finally {
            close();
        }
    }
}
