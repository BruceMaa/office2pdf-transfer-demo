package cn.brucemaa.github.transfer.util;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

/**
 * projectName:office2pdf-transfer-demo
 * cn.brucemaa.github.transfer.util
 *
 * @author Bruce Maa
 * @since 2019-01-23.21:02
 */
public class CmdUtils {

    /**
     * 根据文件名称，查找进程，然后关闭进程
     * @param filename  文件名称
     */
    public static void killProcessWithFileName(String filename) {

        String cmd = String.format("tasklist /v | findstr '%s'", filename);
        try {
            Process process = Runtime.getRuntime().exec(cmd);
            InputStream is = process.getInputStream();
            BufferedReader reader = new BufferedReader(new InputStreamReader(is));
            String line;
            while((line = reader.readLine())!= null){
                System.out.println(line);
                if (line.contains(filename)){
                    line = line.substring(line.indexOf("pid:"));
                    String[] sl = line.split(" ");
                }
            }
            process.waitFor();
            is.close();
            reader.close();
            process.destroy();
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        killProcessWithFileName("aaaaaaaa.doc");
    }
}
