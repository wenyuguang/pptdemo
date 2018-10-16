package sound.extraction;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.hslf.usermodel.SoundData;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * HSLF抽取音频示例
 * Created by cx on 2015/7/1.
 */
public class HSLF4SoundExtraction {
    private static final Logger logger = LoggerFactory.getLogger(HSLF4SoundExtraction.class);

    public static void main(String[] args) throws IOException {
        SlideShow ppt = new SlideShow(new HSLFSlideShow("data/sound_extraction.ppt"));
        SoundData[] sound = ppt.getSoundData();

        for (int i = 0; i < sound.length; i++) {
            if(sound[i].getSoundType().equals(".WAV")){
                //抽取出音频文件输出到指定目录
                FileOutputStream out = new FileOutputStream("output/sound/" + sound[i].getSoundName());
                out.write(sound[i].getData());
                out.close();
            }

            logger.debug("sound file type is {}, file name is {}", sound[i].getSoundType(), sound[i].getSoundName());
        }
    }
}
