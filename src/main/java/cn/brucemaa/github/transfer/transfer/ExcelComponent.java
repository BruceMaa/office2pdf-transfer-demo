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
 * @since 2019-01-23.14:35
 */
public class ExcelComponent {
    private final Logger logger = LoggerFactory.getLogger(this.getClass());

    private ActiveXComponent objExcel;

    private Dispatch wordbook;

    private String filename;

    /**
     * 打开Excel表格
     *
     * @param filename Excel表格路径
     */
    private void open(String filename) {
        this.filename = filename;

        // 初始化com线程
        ComThread.InitMTA(true);

        objExcel = new ActiveXComponent("Excel.Application");
        // 不可见打开Excel
        objExcel.setProperty("Visible", new Variant(false));
        Dispatch workbooks = objExcel.getProperty("Workbooks").toDispatch();
        try {
            wordbook = Dispatch.call(workbooks, "Open", filename).toDispatch();
        } catch (ComFailException e) {
            logger.error("excel filename: {}, open error: {}", filename, e);
            throw e;
        } catch (Exception e) {
            logger.error("excel filename: {}, open error: {}", filename, e);
            throw e;
        }
    }

    /**
     * 将Excel表格另存为pdf文件
     *
     * @param pdfFilename pdf文件地址
     */
    private void saveAsPdf(String pdfFilename) {
        try {
            // 另存储为PDF文件
            Variant variant = Dispatch.call(wordbook, "ExportAsFixedFormat", 0, pdfFilename);
            logger.info("excel filename: {}, pdfFilename: {}, saveAs: {}", filename, pdfFilename, variant);
        } catch (Exception e) {
            logger.error("excel filename: {}, saveAs error: {}", filename, e);
            throw e;
        }
    }

    /**
     * 将Excel表格关闭
     */
    private void close() {
        try {
            Variant variant = Dispatch.call(wordbook, "Close", false);
            logger.info("excel filename: {}, Close: {}", filename, variant);
            variant = objExcel.invoke("Quit", new Variant[0]);
            logger.info("excel filename: {}, Quit: {}", filename, variant);
            variant.safeRelease();
        } catch (Exception e) {
            logger.error("excel filename: {}, Close error: {}", filename, e);
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
            logger.error("excel filename: {}, open error: {}", filename, e);
            throw e;
        } catch (Exception e) {
            logger.error("ExcelComponent saveAsPdf error: {}", e);
            throw e;
        } finally {
            close();
        }
    }
}
