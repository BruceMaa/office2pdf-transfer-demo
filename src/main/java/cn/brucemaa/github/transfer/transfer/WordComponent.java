package cn.brucemaa.github.transfer.transfer;

import cn.brucemaa.github.transfer.util.CmdUtils;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComFailException;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Timer;
import java.util.TimerTask;

/**
 * projectName:office2pdf-transfer-demo
 * cn.brucemaa.github.transfer.transfer
 *
 * @author Bruce Maa
 * @since 2019-01-23.14:34
 */
public class WordComponent {
    private final Logger logger = LoggerFactory.getLogger(this.getClass());
    private ActiveXComponent objWord;

    private Dispatch document;

    private String filename;

    /**
     * 打开word文档
     *
     * @param filename word文档路径
     */
    private void open(String filename) {
        this.filename = filename;

        // 初始化com线程
        ComThread.InitMTA(true);

        objWord = new ActiveXComponent("Word.Application");
        // 不可见打开word
        objWord.setProperty("Visible", new Variant(false));

        // 禁用宏
        objWord.setProperty("AutomationSecurity", new Variant(3));
        objWord.setProperty("DisplayAlerts", new Variant(0));
        objWord.setProperty("ScreenUpdating", new Variant(false));
        Dispatch documents = objWord.getProperty("Documents").toDispatch();
        try {
            Timer timer = new Timer(true);
            TimerTask task = new TimerTask() {
                @Override
                public void run() {
                    System.out.println("===========time is up=========");
                    CmdUtils.killProcessWithFileName(filename);
                }
            };

            timer.schedule(task, 30000);

            document = Dispatch.call(documents, "Open", filename).toDispatch();
        } catch (ComFailException e) {
            logger.error("word filename: {}, open error: {}", filename, e);
            throw e;
        } catch (Exception e) {
            logger.error("word filename: {}, open error: {}", filename, e);
            throw e;
        }
    }

    /**
     * 将word文档另存为pdf文件
     *
     * @param pdfFilename pdf文件地址
     */
    private void saveAsPdf(String pdfFilename) {
        // 另存储为PDF文件
        try {
            Variant variant = Dispatch.call(document, "ExportAsFixedFormat", pdfFilename, 17);
            logger.info("filename: {}, pdfFilename: {}, ExportAsFixedFormat: {}", filename, pdfFilename, variant);
        } catch (Exception e) {
            logger.error("filename: {}, saveAs error: {}", filename, e);
            throw e;
        }
    }

    /**
     * 将word文档关闭
     */
    private void close() {
        try {
            // 最后关闭该文档，并且不保存  false关闭不保存  true关闭并保存，设置为false，可以避免并存为界面卡死
            Variant variant = Dispatch.call(document, "Close", false);
            logger.info("word filename: {}, Close: {}", filename, variant);
            variant = objWord.invoke("Quit", new Variant[0]);
            logger.info("word filename: {}, Quit: {}", filename, variant);
            variant.safeRelease();
        } catch (Exception e) {
            logger.error("word filename: {}, Close error: {}", filename, e);
            throw e;
        } finally {
            // 释放com线程
            ComThread.Release();
        }
    }

    public void saveAsPdf(String sourceFilename, String pdfFileName) {
        try {
            open(sourceFilename);
            saveAsPdf(pdfFileName);
        } catch (ComFailException e) {
            logger.error("word filename: {}, open error: {}", filename, e);
            throw e;
        } catch (Exception e) {
            logger.error("WordComponent saveAsPdf error: {}", e);
            throw e;
        } finally {
            close();
        }
    }

}
