package com.legao.word.poi.vo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.InputStream;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class WordPictureVO {
    private InputStream pictureData;
    private Integer pictureType;
    private String filename;
    private Integer width;
    private Integer height;
}
