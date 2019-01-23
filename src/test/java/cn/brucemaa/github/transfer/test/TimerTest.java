package cn.brucemaa.github.transfer.test;

import org.junit.Test;

import java.util.Timer;
import java.util.TimerTask;

/**
 * projectName:office2pdf-transfer-demo
 * cn.brucemaa.github.transfer.test
 *
 * @author Bruce Maa
 * @since 2019-01-23.19:30
 */
public class TimerTest {

    @Test
    public void test1() {
        Timer timer = new Timer(true);

        TimerTask task = new TimerTask() {
            @Override
            public void run() {
                System.out.println("done");
            }
        };

        timer.schedule(task, 1000);

        try {
            Thread.sleep(2000L);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

}
